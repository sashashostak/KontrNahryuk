/**
 * PythonExcelService - Сервіс для виклику Python скрипта обробки Excel
 */

import { spawn } from 'child_process';
import { join } from 'path';
import { app } from 'electron';
import { writeFileSync, unlinkSync } from 'fs';
import { tmpdir } from 'os';

export interface PythonProcessConfig {
  destinationFile: string;
  sourceFiles: string[];
  sheets: Array<{
    name: string;
    key_column: string;
    data_columns: string[];
    blacklist: string[];
  }>;
}

export interface PythonProcessResult {
  success: boolean;
  total_rows?: number;
  error?: string;
}

export class PythonExcelService {
  
  /**
   * Знайти Python executable
   */
  private static getPythonCommand(): string {
    // В production використовуємо упакований Python
    if (app.isPackaged) {
      const pythonPath = join(process.resourcesPath, 'python', 'python.exe');
      return pythonPath;
    }
    
    // В development використовуємо встановлений Python
    // Спочатку пробуємо знайти в стандартних місцях
    const pythonPaths = [
      'C:\\Users\\sasha\\AppData\\Local\\Programs\\Python\\Python311\\python.exe',
      'C:\\Python311\\python.exe',
      'python'  // fallback на системний PATH
    ];
    
    return pythonPaths[0];  // Використовуємо знайдений Python 3.11
  }
  
  /**
   * Знайти шлях до Python скрипта
   */
  private static getScriptPath(): string {
    if (app.isPackaged) {
      return join(process.resourcesPath, 'python', 'excel_processor.py');
    }
    
    // В development: __dirname = dist/electron/services/
    // Треба піти до кореня проекту: ../../../python/excel_processor.py
    return join(__dirname, '..', '..', '..', 'python', 'excel_processor.py');
  }
  
  /**
   * Обробити Excel файли через Python
   */
  static async processExcel(config: PythonProcessConfig): Promise<PythonProcessResult> {
    return new Promise((resolve, reject) => {
      const pythonCmd = this.getPythonCommand();
      const scriptPath = this.getScriptPath();
      
      // Створюємо тимчасовий JSON файл з конфігурацією (щоб уникнути проблем з кодуванням stdin)
      const tempConfigPath = join(tmpdir(), `stroiovka_config_${Date.now()}.json`);
      
      try {
        // Записуємо конфігурацію у файл з UTF-8 BOM для надійності
        const configJson = JSON.stringify(config, null, 2);
        writeFileSync(tempConfigPath, '\uFEFF' + configJson, 'utf-8');
        
        console.log(`🐍 Запуск Python процесора...`);
        console.log(`   Python: ${pythonCmd}`);
        console.log(`   Скрипт: ${scriptPath}`);
        console.log(`   Config: ${tempConfigPath}`);
        
        // Запускаємо Python процес з шляхом до конфігурації як аргумент
        const pythonProcess = spawn(pythonCmd, [scriptPath, tempConfigPath], {
          stdio: ['ignore', 'pipe', 'pipe'],
          env: {
            ...process.env,
            PYTHONIOENCODING: 'utf-8'
          }
        });
        
        let stdoutData = '';
        let stderrData = '';
        
        // Збираємо вихід
        pythonProcess.stdout.on('data', (data) => {
          const text = data.toString('utf-8');
          stdoutData += text;
          console.log(text.trim());
        });
        
        pythonProcess.stderr.on('data', (data) => {
          const text = data.toString('utf-8');
          stderrData += text;
          console.error(text.trim());
        });
        
        // Обробка завершення
        pythonProcess.on('close', (code) => {
          // Видаляємо тимчасовий файл
          try {
            unlinkSync(tempConfigPath);
          } catch (e) {
            // Ігноруємо помилки видалення
          }
          
          console.log(`🐍 Python процес завершено з кодом: ${code}`);
          
          if (code !== 0) {
            reject(new Error(`Python process exited with code ${code}\n${stderrData}`));
            return;
          }
          
          // Парсимо результат з виходу
          try {
            const resultMatch = stdoutData.match(/__RESULT__(.+?)__END__/s);
            
            if (resultMatch) {
              const result: PythonProcessResult = JSON.parse(resultMatch[1]);
              
              if (result.success) {
                resolve(result);
              } else {
                reject(new Error(result.error || 'Unknown Python error'));
              }
            } else {
              reject(new Error('Could not parse Python result'));
            }
          } catch (error) {
            reject(new Error(`Error parsing Python output: ${error}`));
          }
        });
        
        pythonProcess.on('error', (error) => {
          // Видаляємо тимчасовий файл
          try {
            unlinkSync(tempConfigPath);
          } catch (e) {
            // Ігноруємо помилки видалення
          }
          reject(new Error(`Failed to start Python process: ${error.message}`));
        });
        
      } catch (error) {
        // Видаляємо тимчасовий файл у разі помилки
        try {
          unlinkSync(tempConfigPath);
        } catch (e) {
          // Ігноруємо помилки видалення
        }
        reject(new Error(`Failed to create config file: ${error}`));
      }
    });
  }
}

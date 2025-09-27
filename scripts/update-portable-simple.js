const { spawn } = require('child_process');
const fs = require('fs').promises;
const path = require('path');

async function updatePortableSimple() {
    try {
        console.log('🔄 Простий апдейт Portable версії...');
        
        // Перевіряємо чи існує папка KontrNahryuk-Portable
        const portablePath = path.join(__dirname, '../KontrNahryuk-Portable');
        
        console.log('🔨 Збираємо проект...');
        
        // Збираємо проект
        const buildProcess = spawn('npm', ['run', 'build'], {
            stdio: 'inherit',
            shell: true,
            cwd: process.cwd()
        });
        
        await new Promise((resolve, reject) => {
            buildProcess.on('close', (code) => {
                if (code === 0) {
                    console.log('✅ Збірка завершена успішно');
                    resolve(true);
                } else {
                    console.log(`❌ Збірка завершилася з кодом ${code}`);
                    reject(new Error(`Build failed with code ${code}`));
                }
            });
        });
        
        console.log('📦 Копіюємо скомпільовані файли...');
        
        // Копіюємо збірку до KontrNahryuk-Portable/resources/app
        const targetAppPath = path.join(portablePath, 'resources', 'app');
        
        try {
            await fs.mkdir(targetAppPath, { recursive: true });
            
            // Копіюємо dist папку (скомпільовані файли)
            const distSource = path.join(__dirname, '../dist');
            const distTarget = path.join(targetAppPath, 'dist');
            
            await copyDirectory(distSource, distTarget);
            
            // Копіюємо electron папку (сирцеві TypeScript файли)
            const electronSource = path.join(__dirname, '../electron');
            const electronTarget = path.join(targetAppPath, 'electron');
            
            await copyDirectory(electronSource, electronTarget);
            
            // Копіюємо src папку 
            const srcSource = path.join(__dirname, '../src');
            const srcTarget = path.join(targetAppPath, 'src');
            
            await copyDirectory(srcSource, srcTarget);
            
            // Копіюємо package.json
            const packageSource = path.join(__dirname, '../package.json');
            const packageTarget = path.join(targetAppPath, 'package.json');
            await fs.copyFile(packageSource, packageTarget);
            
            console.log('✅ Файли скопійовано до Portable версії');
            console.log(`📁 Шлях до Portable EXE: ${portablePath}\\KontrNahryuk.exe`);
            
            // Перевіряємо чи є EXE файл
            const exePath = path.join(portablePath, 'KontrNahryuk.exe');
            try {
                await fs.access(exePath);
                console.log('🎯 EXE файл знайдено та готовий до використання!');
                
                // Отримуємо розмір файлу
                const stats = await fs.stat(exePath);
                console.log(`📊 Розмір: ${(stats.size / 1024 / 1024).toFixed(2)} MB`);
                
                // Автоматично змінюємо іконку на свинку
                console.log('🐷 Оновлюємо іконку на свинку...');
                const iconPath = path.join(__dirname, '../build/icon.ico');
                const rceditPath = path.join(__dirname, '../node_modules/rcedit/bin/rcedit.exe');
                
                try {
                    const rceditProcess = spawn(rceditPath, [exePath, '--set-icon', iconPath], {
                        stdio: 'inherit'
                    });
                    
                    await new Promise((resolve, reject) => {
                        rceditProcess.on('close', (code) => {
                            if (code === 0) {
                                console.log('✅ Іконка свинки встановлена!');
                                resolve(true);
                            } else {
                                console.log(`⚠️ Помилка зміни іконки (код ${code})`);
                                resolve(false);
                            }
                        });
                        
                        rceditProcess.on('error', (error) => {
                            console.log('⚠️ Не вдалося змінити іконку:', error.message);
                            resolve(false);
                        });
                    });
                } catch (iconError) {
                    console.log('⚠️ Помилка з іконкою:', iconError.message);
                }
                
            } catch {
                console.log('⚠️ EXE файл не знайдено. Можливо потрібно створити початкову Portable версію вручну.');
            }
            
        } catch (error) {
            console.error('❌ Помилка копіювання файлів:', error);
            throw error;
        }
        
        console.log('🎉 Оновлення Portable версії завершено!');
        
    } catch (error) {
        console.error('❌ Помилка оновлення Portable версії:', error);
        process.exit(1);
    }
}

// Функція для рекурсивного копіювання директорій
async function copyDirectory(source, target) {
    try {
        await fs.mkdir(target, { recursive: true });
        const files = await fs.readdir(source);
        
        for (const file of files) {
            const sourcePath = path.join(source, file);
            const targetPath = path.join(target, file);
            
            const stat = await fs.stat(sourcePath);
            if (stat.isFile()) {
                await fs.copyFile(sourcePath, targetPath);
            } else if (stat.isDirectory()) {
                await copyDirectory(sourcePath, targetPath);
            }
        }
    } catch (error) {
        console.warn(`⚠️ Не вдалося скопіювати ${source}:`, error.message);
    }
}

// Запуск скрипту
if (require.main === module) {
    updatePortableSimple();
}

module.exports = { updatePortableSimple };
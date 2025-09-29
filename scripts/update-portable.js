const { spawn } = require('child_process');
const fs = require('fs').promises;
const path = require('path');

async function updatePortableVersion() {
    try {
        console.log('🔄 Автоматичне оновлення Portable версії...');
        
        // Перевіряємо чи існує папка KontrNahryuk-Portable
        const portablePath = path.join(__dirname, '../KontrNahryuk-Portable');
        
        try {
            await fs.access(portablePath);
            console.log('📁 Папка KontrNahryuk-Portable знайдена');
        } catch {
            console.log('❌ Папка KontrNahryuk-Portable не знайдена, створюємо...');
            await fs.mkdir(portablePath, { recursive: true });
        }
        
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
        
        console.log('📦 Створюємо Portable версію...');
        
        // Створюємо Portable версію через electron-builder
        const packageProcess = spawn('npm', ['run', 'dist:portable'], {
            stdio: 'inherit',
            shell: true,
            cwd: process.cwd()
        });
        
        await new Promise((resolve, reject) => {
            packageProcess.on('close', (code) => {
                if (code === 0) {
                    console.log('✅ Portable версія створена');
                    resolve(true);
                } else {
                    console.log(`❌ Packaging завершився з кодом ${code}`);
                    reject(new Error(`Packaging failed with code ${code}`));
                }
            });
        });
        
        // Копіюємо файли з release до KontrNahryuk-Portable
        const releasePath = path.join(__dirname, '../release');
        
        // Перевіряємо чи існує .exe файл напряму в release
        let sourceFile = null;
        try {
            const releaseFiles = await fs.readdir(releasePath);
            const exeFile = releaseFiles.find(f => f.endsWith('.exe'));
            if (exeFile) {
                sourceFile = path.join(releasePath, exeFile);
                console.log(`🎯 Знайдено Portable EXE: ${exeFile}`);
            }
        } catch (err) {
            console.log('Папка release не знайдена, шукаємо інші варіанти...');
        }
        
        // Якщо не знайшли EXE, шукаємо в підпапках
        let sourceDir = null;
        if (!sourceFile) {
            try {
                const releaseDir = await fs.readdir(releasePath);
                for (const dir of releaseDir) {
                    const dirPath = path.join(releasePath, dir);
                    const stat = await fs.stat(dirPath);
                    if (stat.isDirectory() && (dir.includes('KontrNahryuk') || dir.includes('kontr-nahryuk'))) {
                        sourceDir = dirPath;
                        break;
                    }
                }
            } catch (err) {
                console.warn('Не вдалося знайти release папку:', err.message);
            }
        }
        
        if (sourceFile) {
            console.log('🔄 Копіюємо Portable EXE...');
            const targetPath = path.join(portablePath, 'KontrNahryuk.exe');
            await fs.copyFile(sourceFile, targetPath);
            console.log('✅ Portable EXE файл скопійовано');
        } else if (sourceDir) {
            console.log('🔄 Копіюємо файли до KontrNahryuk-Portable...');
            
            // Видаляємо старі файли
            const portableFiles = await fs.readdir(portablePath).catch(() => []);
            for (const file of portableFiles) {
                const filePath = path.join(portablePath, file);
                const stat = await fs.stat(filePath).catch(() => null);
                if (stat?.isFile()) {
                    await fs.unlink(filePath);
                }
            }
            
            // Копіюємо нові файли
            const sourceFiles = await fs.readdir(sourceDir);
            for (const file of sourceFiles) {
                const sourcePath = path.join(sourceDir, file);
                const targetPath = path.join(portablePath, file);
                
                const stat = await fs.stat(sourcePath);
                if (stat.isFile()) {
                    await fs.copyFile(sourcePath, targetPath);
                } else if (stat.isDirectory()) {
                    await copyDirectory(sourcePath, targetPath);
                }
            }
            
            console.log('✅ Файли успішно скопійовано');
            
            // Перевіряємо чи є exe файл
            const exeFiles = sourceFiles.filter(f => f.endsWith('.exe'));
            if (exeFiles.length > 0) {
                console.log(`🎯 EXE файл знайдено: ${exeFiles[0]}`);
                console.log(`📁 Шлях: ${portablePath}\\${exeFiles[0]}`);
            }
            
        } else {
            console.log('❌ Не знайдено папку з executable файлами в release');
        }
        
        console.log('🎉 Оновлення Portable версії завершено!');
        
    } catch (error) {
        console.error('❌ Помилка оновлення Portable версії:', error);
        process.exit(1);
    }
}

// Функція для рекурсивного копіювання директорій
async function copyDirectory(source, target) {
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
}

// Запуск скрипту
if (require.main === module) {
    updatePortableVersion();
}

module.exports = { updatePortableVersion };
const { spawn } = require('child_process');
const path = require('path');

async function forceUpdateIcon() {
    try {
        console.log('🔄 Примусове оновлення іконки EXE файлу...');
        
        const exePath = path.join(__dirname, '../release/KontrNahryuk-win32-x64/KontrNahryuk.exe');
        const iconPath = path.join(__dirname, '../build/icon.ico');
        const rceditPath = path.join(__dirname, '../node_modules/rcedit/bin/rcedit.exe');
        
        // 1. Встановлюємо іконку з додатковими версійними данними
        console.log('🐷 Встановлюємо іконку з метаданими...');
        
        const rceditProcess = spawn(rceditPath, [
            exePath,
            '--set-icon', iconPath,
            '--set-version-string', 'CompanyName', 'KontrNahryuk Team',
            '--set-version-string', 'FileDescription', 'Ukrainian Document Processor 🐷',
            '--set-version-string', 'ProductName', 'KontrNahryuk',
            '--set-version-string', 'ProductVersion', '1.1.2',
            '--set-version-string', 'FileVersion', '1.1.2',
            '--set-version-string', 'LegalCopyright', 'Copyright © 2024-2025 KontrNahryuk',
            '--set-version-string', 'OriginalFilename', 'KontrNahryuk.exe'
        ], {
            stdio: 'inherit'
        });
        
        await new Promise((resolve, reject) => {
            rceditProcess.on('close', (code) => {
                if (code === 0) {
                    console.log('✅ Іконка та метадані встановлені!');
                    resolve(true);
                } else {
                    console.log(`❌ Помилка (код ${code})`);
                    reject(new Error(`rcedit failed with code ${code}`));
                }
            });
        });
        
        console.log('🔄 Очищуємо кеш іконок Windows...');
        
        // 2. Очищуємо кеш іконок
        const commands = [
            'ie4uinit.exe -show',
            'ie4uinit.exe -ClearIconCache',
            'taskkill /f /im explorer.exe',
            'timeout /t 2 /nobreak > nul',
            'start explorer.exe'
        ];
        
        for (const cmd of commands) {
            try {
                console.log(`🔧 Виконуємо: ${cmd}`);
                const [command, ...args] = cmd.split(' ');
                
                if (command === 'timeout' || command === 'start') {
                    // Для цих команд використовуємо shell
                    await new Promise((resolve) => {
                        const proc = spawn('cmd', ['/c', cmd], { stdio: 'inherit' });
                        proc.on('close', () => resolve(true));
                    });
                } else {
                    await new Promise((resolve) => {
                        const proc = spawn(command, args, { stdio: 'inherit' });
                        proc.on('close', () => resolve(true));
                        proc.on('error', () => resolve(true)); // Ігноруємо помилки
                    });
                }
            } catch (error) {
                console.log(`⚠️ Команда ${cmd} не виконалася, продовжуємо...`);
            }
        }
        
        console.log('✅ Оновлення іконки завершено!');
        console.log('💡 Якщо іконка все ще не змінилася:');
        console.log('   1. Перезавантажте File Explorer');
        console.log('   2. Або перезавантажте комп\'ютер');
        console.log('   3. Спробуйте видалити і заново скопіювати EXE файл');
        
    } catch (error) {
        console.error('❌ Помилка оновлення іконки:', error);
    }
}

if (require.main === module) {
    forceUpdateIcon();
}
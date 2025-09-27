const fs = require('fs');
const path = require('path');
const { spawn } = require('child_process');

async function changeExeIcon() {
    try {
        console.log('🐷 Змінюємо іконку EXE файлу на свинку...');
        
        const exePath = path.join(__dirname, '../KontrNahryuk-Portable/KontrNahryuk.exe');
        const iconPath = path.join(__dirname, '../build/icon.ico');
        const tempExePath = path.join(__dirname, '../KontrNahryuk-Portable/KontrNahryuk_temp.exe');
        
        // Перевіряємо чи існують файли
        if (!fs.existsSync(exePath)) {
            console.log('❌ EXE файл не знайдено:', exePath);
            return false;
        }
        
        if (!fs.existsSync(iconPath)) {
            console.log('❌ ICO файл не знайдено:', iconPath);
            return false;
        }
        
        console.log('📁 EXE файл:', exePath);
        console.log('🐷 Іконка:', iconPath);
        
        // Спробуємо використовути rcedit якщо він встановлений
        try {
            const rceditProcess = spawn('npx', ['rcedit', exePath, '--set-icon', iconPath], {
                stdio: 'inherit',
                shell: true
            });
            
            await new Promise((resolve, reject) => {
                rceditProcess.on('close', (code) => {
                    if (code === 0) {
                        console.log('✅ Іконка успішно змінена через rcedit!');
                        resolve(true);
                    } else {
                        console.log(`⚠️ rcedit завершився з кодом ${code}, спробуємо інший метод`);
                        resolve(false);
                    }
                });
                
                rceditProcess.on('error', (error) => {
                    console.log('⚠️ rcedit недоступний, спробуємо інший метод');
                    resolve(false);
                });
            });
            
            return true;
            
        } catch (error) {
            console.log('⚠️ Помилка з rcedit:', error.message);
        }
        
        // Якщо rcedit не працює, просто повідомимо користувача
        console.log('📝 Для зміни іконки EXE файлу потрібно:');
        console.log('   1. Встановити rcedit: npm install -g rcedit');
        console.log('   2. Або використати ResourceHacker');
        console.log('   3. Або пересоборити EXE з новою іконкою');
        
        return false;
        
    } catch (error) {
        console.error('❌ Помилка зміни іконки:', error);
        return false;
    }
}

// Запуск скрипту
if (require.main === module) {
    changeExeIcon();
}

module.exports = { changeExeIcon };
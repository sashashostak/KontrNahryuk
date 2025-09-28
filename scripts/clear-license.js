const fs = require('fs').promises;
const path = require('path');
const os = require('os');

async function clearLicenseData() {
    try {
        const appDataPath = path.join(os.homedir(), 'AppData', 'Roaming', 'kontr-nahryuk', 'db.json');
        
        console.log('🗑️ Очищення збережених даних ліцензії...');
        console.log('📁 Шлях:', appDataPath);
        
        try {
            const data = await fs.readFile(appDataPath, 'utf8');
            console.log('📄 Поточний вміст db.json:');
            console.log(data);
            
            // Очищуємо licenseKey але залишаємо інші налаштування
            const jsonData = JSON.parse(data);
            if (jsonData.settings && jsonData.settings.licenseKey) {
                delete jsonData.settings.licenseKey;
                console.log('🔑 Видаляємо licenseKey з налаштувань...');
            } else {
                console.log('⚠️ licenseKey не знайдено в налаштуваннях');
            }
            
            await fs.writeFile(appDataPath, JSON.stringify(jsonData, null, 2), 'utf8');
            console.log('✅ Дані очищено, licenseKey видалено');
            
        } catch (error) {
            if (error.code === 'ENOENT') {
                console.log('⚠️ Файл db.json не існує (це нормально для першого запуску)');
            } else {
                console.error('❌ Помилка читання/запису файлу:', error);
            }
        }
        
    } catch (error) {
        console.error('❌ Помилка очищення даних:', error);
    }
}

if (require.main === module) {
    clearLicenseData();
}
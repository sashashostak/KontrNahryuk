const fs = require('fs');
const path = require('path');
const { updatePortableVersion } = require('./update-portable.js');

let isUpdating = false;
let updateQueue = false;

async function watchAndUpdate() {
    console.log('👀 Відстеження змін в коді для автооновлення Portable версії...');
    
    // Папки для відстеження
    const watchDirectories = [
        path.join(__dirname, '../src'),
        path.join(__dirname, '../electron'),
        path.join(__dirname, '../build')
    ];
    
    // Функція для обробки змін
    const handleChange = async (eventType, filename) => {
        if (isUpdating) {
            updateQueue = true;
            return;
        }
        
        console.log(`📝 Зміна виявлена: ${filename} (${eventType})`);
        
        // Затримка для уникнення множинних оновлень
        setTimeout(async () => {
            if (updateQueue) {
                updateQueue = false;
                return;
            }
            
            isUpdating = true;
            try {
                console.log('🔄 Запуск автооновлення Portable версії...');
                await updatePortableVersion();
                console.log('✅ Portable версія оновлена автоматично');
            } catch (error) {
                console.error('❌ Помилка автооновлення:', error);
            } finally {
                isUpdating = false;
                
                // Якщо були додаткові зміни під час оновлення
                if (updateQueue) {
                    updateQueue = false;
                    setTimeout(() => handleChange('change', 'queued'), 1000);
                }
            }
        }, 2000); // 2 секунди затримки
    };
    
    // Налаштовуємо відстеження для кожної папки
    watchDirectories.forEach(dir => {
        try {
            fs.watch(dir, { recursive: true }, handleChange);
            console.log(`👁️ Відстеження: ${dir}`);
        } catch (error) {
            console.warn(`⚠️ Не вдалося налаштувати відстеження для ${dir}:`, error.message);
        }
    });
    
    console.log('🎯 Автооновлення Portable версії налаштовано!');
    console.log('💡 Тепер при зміні файлів Portable версія буде оновлюватися автоматично');
    
    // Зробимо початкове оновлення
    try {
        await updatePortableVersion();
        console.log('🚀 Початкове оновлення завершено');
    } catch (error) {
        console.error('❌ Помилка початкового оновлення:', error);
    }
}

// Запуск якщо скрипт викликано напряму
if (require.main === module) {
    watchAndUpdate().catch(console.error);
}

module.exports = { watchAndUpdate };
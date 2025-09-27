const sharp = require('sharp');
const fs = require('fs');
const path = require('path');

async function convertSvgToIco() {
    try {
        console.log('🔄 Конвертація SVG в ICO...');
        
        // Читаємо SVG файл
        const svgPath = path.join(__dirname, '../build/pig-icon.svg');
        const icoPath = path.join(__dirname, '../build/icon.ico');
        
        // Конвертуємо SVG в PNG 256x256, потім в ICO
        await sharp(svgPath)
            .resize(256, 256)
            .png()
            .toFile(path.join(__dirname, '../build/icon-256.png'));
            
        console.log('✅ PNG файл створено');
        
        // Створюємо різні розміри для ICO
        const sizes = [16, 32, 48, 64, 128, 256];
        
        for (const size of sizes) {
            await sharp(svgPath)
                .resize(size, size)
                .png()
                .toFile(path.join(__dirname, `../build/icon-${size}.png`));
        }
        
        console.log('✅ Всі розміри створено');
        console.log('📝 Тепер потрібно вручну створити ICO файл з PNG файлів');
        console.log('💡 Рекомендую використати онлайн сервіс або ImageMagick');
        
        // Створимо простий ICO з основного PNG
        await sharp(svgPath)
            .resize(256, 256)
            .png()
            .toFile(icoPath.replace('.ico', '.png'));
            
        console.log('✅ Основний PNG для ICO готовий');
        
    } catch (error) {
        console.error('❌ Помилка конвертації:', error);
    }
}

convertSvgToIco();
const sharp = require('sharp');
const png2icons = require('png2icons');
const fs = require('fs').promises;
const path = require('path');

async function createIcoFromSvg() {
    try {
        console.log('🐷 Створюємо іконку свинки для EXE...');
        
        const svgPath = path.join(__dirname, '../build/pig-icon.svg');
        const icoPath = path.join(__dirname, '../build/icon.ico');
        
        console.log('🔄 Конвертуємо SVG в ICO...');
        
        // Створюємо PNG зображення 256x256 для конвертації в ICO
        const pngBuffer = await sharp(svgPath)
            .resize(256, 256, {
                kernel: sharp.kernel.lanczos3,
                fit: 'contain', 
                background: { r: 0, g: 0, b: 0, alpha: 0 }
            })
            .png()
            .toBuffer();
        
        // Конвертуємо PNG в правильний ICO файл
        const icoBuffer = png2icons.createICO(pngBuffer, png2icons.BILINEAR, 0, false, true);
        
        // Зберігаємо ICO файл
        await fs.writeFile(icoPath, icoBuffer);
        
        // Також створюємо додаткові розміри для різних цілей
        const sizes = [16, 32, 48, 64, 128];
        for (const size of sizes) {
            const sizeBuffer = await sharp(svgPath)
                .resize(size, size, {
                    kernel: sharp.kernel.lanczos3,
                    fit: 'contain',
                    background: { r: 0, g: 0, b: 0, alpha: 0 }
                })
                .png()
                .toBuffer();
            
            await fs.writeFile(
                path.join(__dirname, `../build/icon-${size}.png`),
                sizeBuffer
            );
        }
        
        console.log('✅ ICO файл створено:', icoPath);
        console.log('🐷 Іконка свинки готова!');
        
    } catch (error) {
        console.error('❌ Помилка створення ICO:', error);
    }
}

createIcoFromSvg();
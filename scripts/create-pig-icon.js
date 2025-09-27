const sharp = require('sharp');
const toIco = require('to-ico');
const fs = require('fs').promises;
const path = require('path');

async function createIcoFromSvg() {
    try {
        console.log('🐷 Створюємо іконку свинки для EXE...');
        
        const svgPath = path.join(__dirname, '../build/pig-icon.svg');
        const icoPath = path.join(__dirname, '../build/icon.ico');
        
        // Створюємо PNG файли різних розмірів
        const sizes = [16, 24, 32, 48, 64, 128, 256];
        const pngBuffers = [];
        
        for (const size of sizes) {
            console.log(`📐 Створюємо розмір ${size}x${size}...`);
            const buffer = await sharp(svgPath)
                .resize(size, size, {
                    kernel: sharp.kernel.lanczos3,
                    fit: 'contain',
                    background: { r: 0, g: 0, b: 0, alpha: 0 }
                })
                .png()
                .toBuffer();
            
            pngBuffers.push(buffer);
        }
        
        console.log('🔄 Конвертуємо PNG в ICO...');
        
        // Створюємо ICO файл
        const icoBuffer = await toIco(pngBuffers);
        
        // Зберігаємо ICO файл
        await fs.writeFile(icoPath, icoBuffer);
        
        console.log('✅ ICO файл створено:', icoPath);
        console.log('🐷 Іконка свинки готова!');
        
    } catch (error) {
        console.error('❌ Помилка створення ICO:', error);
    }
}

createIcoFromSvg();
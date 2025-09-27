const sharp = require('sharp');
const fs = require('fs').promises;
const path = require('path');

async function createSimpleIcon() {
    try {
        console.log('🐷 Створюємо просту іконку свинки...');
        
        const svgPath = path.join(__dirname, '../build/simple-pig-icon.svg');
        const icoPath = path.join(__dirname, '../build/simple-icon.ico');
        
        // Створюємо PNG розміром 256x256
        const pngBuffer = await sharp(svgPath)
            .resize(256, 256, {
                kernel: sharp.kernel.lanczos3,
                fit: 'contain',
                background: { r: 255, g: 182, b: 193, alpha: 1 }
            })
            .png()
            .toBuffer();
        
        // Зберігаємо як PNG для тестування
        await fs.writeFile(path.join(__dirname, '../build/simple-icon.png'), pngBuffer);
        
        // Створюємо ICO файл вручну (Windows ICO format)
        const iconHeader = Buffer.alloc(22);
        iconHeader.writeUInt16LE(0, 0);        // Reserved
        iconHeader.writeUInt16LE(1, 2);        // Type: 1 for ICO
        iconHeader.writeUInt16LE(1, 4);        // Number of images
        
        // Image directory entry
        iconHeader.writeUInt8(0, 6);           // Width (0 = 256)
        iconHeader.writeUInt8(0, 7);           // Height (0 = 256) 
        iconHeader.writeUInt8(0, 8);           // Colors (0 = no palette)
        iconHeader.writeUInt8(0, 9);           // Reserved
        iconHeader.writeUInt16LE(1, 10);       // Color planes
        iconHeader.writeUInt16LE(32, 12);      // Bits per pixel
        iconHeader.writeUInt32LE(pngBuffer.length, 14); // Image size
        iconHeader.writeUInt32LE(22, 18);      // Offset to image data
        
        // Комбінуємо header і PNG data
        const icoBuffer = Buffer.concat([iconHeader, pngBuffer]);
        
        await fs.writeFile(icoPath, icoBuffer);
        
        console.log('✅ Просту іконку створено:', icoPath);
        
        return icoPath;
        
    } catch (error) {
        console.error('❌ Помилка створення простої іконки:', error);
        throw error;
    }
}

if (require.main === module) {
    createSimpleIcon();
}

module.exports = { createSimpleIcon };
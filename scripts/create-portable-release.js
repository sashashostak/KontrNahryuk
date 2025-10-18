/**
 * Скрипт для автоматичного створення portable релізу KontrNahryuk
 * 
 * Процес:
 * 1. Зчитує поточну версію з package.json
 * 2. Запаковує папку release/KontrNahryuk-win32-x64 у ZIP архів
 * 3. Називає архів: KontrNahryuk-v{version}-portable.zip
 * 4. Зберігає архів у кореневу папку проекту
 * 
 * Використання: node scripts/create-portable-release.js
 */

const fs = require('fs');
const path = require('path');
const archiver = require('archiver');

// Кольоровий вивід для консолі
const colors = {
  reset: '\x1b[0m',
  green: '\x1b[32m',
  blue: '\x1b[34m',
  yellow: '\x1b[33m',
  red: '\x1b[31m',
};

function log(message, color = colors.reset) {
  console.log(`${color}${message}${colors.reset}`);
}

async function createPortableRelease() {
  try {
    // 1. Зчитуємо версію з package.json
    const packageJson = JSON.parse(fs.readFileSync('package.json', 'utf8'));
    const version = packageJson.version;
    
    log(`\n🚀 Створення portable релізу KontrNahryuk v${version}`, colors.blue);
    
    // 2. Перевіряємо чи існує папка зібраного додатку
    const sourcePath = path.join('release', 'KontrNahryuk-win32-x64');
    if (!fs.existsSync(sourcePath)) {
      log('❌ Помилка: Папка release/KontrNahryuk-win32-x64 не знайдена!', colors.red);
      log('   Спочатку запустіть: npm run package', colors.yellow);
      process.exit(1);
    }
    
    // 3. Формуємо назву архіву згідно ТЗ: має містити "portable" та закінчуватись на .zip
    const zipName = `KontrNahryuk-v${version}-portable.zip`;
    const outputPath = path.resolve(zipName);
    
    // 4. Видаляємо старий архів якщо існує
    if (fs.existsSync(outputPath)) {
      log(`📁 Видаляємо старий архів: ${zipName}`, colors.yellow);
      fs.unlinkSync(outputPath);
    }
    
    // 5. Створюємо новий архів
    log(`📦 Створюємо архів: ${zipName}`, colors.blue);
    
    const output = fs.createWriteStream(outputPath);
    const archive = archiver('zip', {
      zlib: { level: 9 } // Максимальне стиснення
    });
    
    // Обробка помилок
    output.on('error', (err) => {
      log(`❌ Помилка запису файлу: ${err.message}`, colors.red);
      process.exit(1);
    });
    
    archive.on('error', (err) => {
      log(`❌ Помилка архівування: ${err.message}`, colors.red);
      process.exit(1);
    });
    
    // Відстежуємо прогрес
    archive.on('progress', (progress) => {
      const percent = ((progress.entries.processed / progress.entries.total) * 100).toFixed(1);
      process.stdout.write(`\r   Прогрес: ${percent}% (${progress.entries.processed}/${progress.entries.total} файлів)`);
    });
    
    // Подія завершення
    output.on('close', () => {
      const sizeInMB = (archive.pointer() / 1024 / 1024).toFixed(2);
      log(`\n✅ Реліз успішно створено!`, colors.green);
      log(`📊 Розмір архіву: ${sizeInMB} MB`, colors.blue);
      log(`📁 Файл: ${zipName}`, colors.blue);
      log(`\n📝 Наступні кроки:`, colors.yellow);
      log(`   1. Перейдіть на: https://github.com/sashashostak/KontrNahryuk/releases/new`);
      log(`   2. Встановіть тег: v${version}`);
      log(`   3. Завантажте файл: ${zipName}`);
      log(`   4. Опублікуйте реліз`);
    });
    
    // Підключаємо вихідний потік до архіватора
    archive.pipe(output);
    
    // Додаємо всі файли з папки до архіву
    // Зберігаємо структуру папок: KontrNahryuk-win32-x64/...
    archive.directory(sourcePath, 'KontrNahryuk-win32-x64');
    
    // Завершуємо архівування
    await archive.finalize();
    
  } catch (error) {
    log(`\n❌ Помилка: ${error.message}`, colors.red);
    console.error(error);
    process.exit(1);
  }
}

// Запуск
createPortableRelease();

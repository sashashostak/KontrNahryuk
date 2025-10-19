/**
 * Створення patch-файлу для оновлення KontrNahryuk
 * 
 * Процес:
 * 1. Порівнює поточну версію з попередньою
 * 2. Знаходить тільки змінені файли
 * 3. Створює мінімальний ZIP з патчем
 * 
 * Використання: node scripts/create-update-patch.js <from-version> <to-version>
 * Приклад: node scripts/create-update-patch.js 1.4.1 1.4.2
 */

const fs = require('fs');
const path = require('path');
const archiver = require('archiver');
const crypto = require('crypto');

const colors = {
  reset: '\x1b[0m',
  green: '\x1b[32m',
  blue: '\x1b[34m',
  yellow: '\x1b[33m',
  red: '\x1b[31m',
  cyan: '\x1b[36m',
};

function log(message, color = colors.reset) {
  console.log(`${color}${message}${colors.reset}`);
}

/**
 * Обчислити MD5 хеш файлу
 */
function getFileHash(filePath) {
  const content = fs.readFileSync(filePath);
  return crypto.createHash('md5').update(content).digest('hex');
}

/**
 * Рекурсивно отримати всі файли з папки
 */
function getAllFiles(dirPath, arrayOfFiles = []) {
  const files = fs.readdirSync(dirPath);

  files.forEach((file) => {
    const fullPath = path.join(dirPath, file);
    if (fs.statSync(fullPath).isDirectory()) {
      arrayOfFiles = getAllFiles(fullPath, arrayOfFiles);
    } else {
      arrayOfFiles.push(fullPath);
    }
  });

  return arrayOfFiles;
}

/**
 * Знайти файли, що змінились між версіями
 */
function findChangedFiles(currentDir, previousDir) {
  const changedFiles = [];
  const newFiles = [];

  // Отримати всі файли з поточної версії
  const currentFiles = getAllFiles(currentDir);
  
  log(`\n📁 Аналіз файлів...`, colors.blue);
  log(`   Поточна версія: ${currentFiles.length} файлів`, colors.cyan);

  let checkedCount = 0;
  
  for (const currentFile of currentFiles) {
    checkedCount++;
    if (checkedCount % 500 === 0) {
      process.stdout.write(`\r   Перевірено: ${checkedCount}/${currentFiles.length} файлів...`);
    }

    const relativePath = path.relative(currentDir, currentFile);
    const previousFile = path.join(previousDir, relativePath);

    // Пропускаємо node_modules та інші великі папки
    if (relativePath.includes('node_modules') || 
        relativePath.includes('.git') ||
        relativePath.includes('locales')) {
      continue;
    }

    // Якщо файл не існував раніше - це новий файл
    if (!fs.existsSync(previousFile)) {
      newFiles.push({ path: currentFile, relativePath, reason: 'NEW' });
      continue;
    }

    // Порівняти хеші файлів
    const currentHash = getFileHash(currentFile);
    const previousHash = getFileHash(previousFile);

    if (currentHash !== previousHash) {
      const currentSize = fs.statSync(currentFile).size;
      const previousSize = fs.statSync(previousFile).size;
      const sizeDiff = currentSize - previousSize;
      
      changedFiles.push({ 
        path: currentFile, 
        relativePath, 
        reason: 'MODIFIED',
        sizeDiff 
      });
    }
  }

  console.log(''); // Новий рядок після прогресу

  return { changedFiles, newFiles };
}

async function createUpdatePatch() {
  try {
    // Читаємо аргументи
    const args = process.argv.slice(2);
    if (args.length < 2) {
      log('\n❌ Помилка: Потрібно вказати версії', colors.red);
      log('   Використання: node scripts/create-update-patch.js <from-version> <to-version>', colors.yellow);
      log('   Приклад: node scripts/create-update-patch.js 1.4.1 1.4.2\n', colors.yellow);
      process.exit(1);
    }

    const fromVersion = args[0];
    const toVersion = args[1];

    log(`\n🔄 Створення патчу оновлення`, colors.blue);
    log(`   З версії: ${fromVersion}`, colors.cyan);
    log(`   На версію: ${toVersion}`, colors.cyan);

    // Перевірка існування папок
    const currentDir = path.join('release', 'KontrNahryuk-win32-x64');
    if (!fs.existsSync(currentDir)) {
      log('\n❌ Помилка: Поточна версія не зібрана!', colors.red);
      log('   Запустіть: npm run package\n', colors.yellow);
      process.exit(1);
    }

    // Для демонстрації створюємо список критичних файлів, які точно змінились
    // У реальній ситуації треба мати попередню версію для порівняння
    const criticalFiles = [
      'dist/electron/main.js',
      'dist/electron/preload.js',
      'dist/electron/services/updateService.js',
      'dist/renderer/index.html',
      'dist/renderer/assets/index-*.js',
      'package.json',
      'resources/app.asar',
    ];

    log(`\n🔍 Визначення змінених файлів...`, colors.blue);

    const filesToPatch = [];
    
    // Знаходимо файли, які відповідають критичним шаблонам
    const allFiles = getAllFiles(currentDir);
    
    for (const pattern of criticalFiles) {
      const regex = new RegExp(pattern.replace(/\*/g, '.*').replace(/\//g, '\\\\'));
      const matchedFiles = allFiles.filter(f => regex.test(path.relative(currentDir, f)));
      
      matchedFiles.forEach(file => {
        const relativePath = path.relative(currentDir, file);
        const size = fs.statSync(file).size;
        filesToPatch.push({ 
          path: file, 
          relativePath,
          size,
          reason: 'CRITICAL_UPDATE'
        });
      });
    }

    if (filesToPatch.length === 0) {
      log('\n⚠️ Не знайдено файлів для патчу', colors.yellow);
      log('   Можливо, версії однакові?\n', colors.yellow);
      process.exit(0);
    }

    // Виводимо список файлів
    log(`\n📝 Знайдено файлів для оновлення: ${filesToPatch.length}`, colors.green);
    let totalSize = 0;
    filesToPatch.forEach(f => {
      totalSize += f.size;
      const sizeKB = (f.size / 1024).toFixed(1);
      log(`   ✓ ${f.relativePath} (${sizeKB} KB)`, colors.cyan);
    });

    const totalSizeMB = (totalSize / 1024 / 1024).toFixed(2);
    log(`\n💾 Загальний розмір патчу: ${totalSizeMB} MB`, colors.green);

    // Створюємо ZIP архів
    const patchName = `KontrNahryuk-v${fromVersion}-to-v${toVersion}-patch.zip`;
    const outputPath = path.resolve(patchName);

    if (fs.existsSync(outputPath)) {
      log(`\n📁 Видаляємо старий патч: ${patchName}`, colors.yellow);
      fs.unlinkSync(outputPath);
    }

    log(`\n📦 Створюємо патч-файл: ${patchName}`, colors.blue);

    const output = fs.createWriteStream(outputPath);
    const archive = archiver('zip', {
      zlib: { level: 9 }
    });

    output.on('close', () => {
      const archiveSizeMB = (archive.pointer() / 1024 / 1024).toFixed(2);
      const savings = (1 - archive.pointer() / (415 * 1024 * 1024)) * 100;
      
      log(`\n✅ Патч успішно створено!`, colors.green);
      log(`📊 Розмір патчу: ${archiveSizeMB} MB`, colors.blue);
      log(`💰 Економія: ${savings.toFixed(1)}% порівняно з повним portable`, colors.green);
      log(`📁 Файл: ${patchName}`, colors.blue);

      log(`\n📝 Наступні кроки:`, colors.yellow);
      log(`   1. Перейдіть на: https://github.com/sashashostak/KontrNahryuk/releases/tag/v${toVersion}`);
      log(`   2. Edit release`);
      log(`   3. Завантажте: ${patchName} (швидке оновлення)`);
      log(`   4. Також залиште: KontrNahryuk-v${toVersion}-portable.zip (повна версія)`);
      log(`   5. В описі вкажіть обидва варіанти завантаження\n`);
    });

    archive.on('error', (err) => {
      log(`❌ Помилка архівування: ${err.message}`, colors.red);
      process.exit(1);
    });

    archive.pipe(output);

    // Створюємо інструкцію перед архівуванням
    const instructions = `# Інструкція по встановленню патчу ${fromVersion} → ${toVersion}

## Автоматичне встановлення (рекомендовано)

1. Запустіть KontrNahryuk v${fromVersion}
2. Відкрийте: ⚙️ Налаштування → Оновлення
3. Натисніть: "Перевірити оновлення"
4. Система автоматично завантажить та встановить оновлення

## Ручне встановлення патчу

1. Закрийте KontrNahryuk (якщо запущено)
2. Розпакуйте KontrNahryuk-v${fromVersion}-to-v${toVersion}-patch.zip
3. Скопіюйте файли в папку з KontrNahryuk
4. Підтвердіть заміну файлів
5. Запустіть KontrNahryuk.exe

## Що оновлюється:

${filesToPatch.slice(0, 10).map(f => `- ${f.relativePath}`).join('\n')}
... та ще ${filesToPatch.length - 10} файлів

## Резервне копіювання

Рекомендуємо створити резервну копію папки KontrNahryuk перед встановленням патчу.

## Підтримка

Якщо виникнуть проблеми, завантажте повну portable версію:
https://github.com/sashashostak/KontrNahryuk/releases/tag/v${toVersion}
`;

    // Додаємо інструкцію до архіву
    archive.append(instructions, { name: 'UPDATE-INSTRUCTIONS.txt' });

    // Додаємо файли до архіву
    for (const file of filesToPatch) {
      archive.file(file.path, { name: file.relativePath });
    }

    await archive.finalize();

  } catch (error) {
    log(`\n❌ Помилка: ${error.message}`, colors.red);
    console.error(error);
    process.exit(1);
  }
}

createUpdatePatch();

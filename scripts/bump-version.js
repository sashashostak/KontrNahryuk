#!/usr/bin/env node

/**
 * Автоматичне оновлення версії проекту KontrNahryuk
 * Інкрементує patch версію (1.1.0 -> 1.1.1) у всіх файлах проекту
 */

const fs = require('fs');
const path = require('path');

// Файли де потрібно оновити версії
const FILES_TO_UPDATE = [
  'package.json',
  'src/main.ts',
  'electron/services/updateService.ts'
];

function getCurrentVersion() {
  const packageJson = JSON.parse(fs.readFileSync('package.json', 'utf8'));
  return packageJson.version;
}

function incrementPatchVersion(version) {
  const [major, minor, patch] = version.split('.').map(Number);
  return `${major}.${minor}.${patch + 1}`;
}

function updatePackageJson(newVersion) {
  const packagePath = 'package.json';
  const packageJson = JSON.parse(fs.readFileSync(packagePath, 'utf8'));
  packageJson.version = newVersion;
  fs.writeFileSync(packagePath, JSON.stringify(packageJson, null, 2) + '\n');
  console.log(`✅ Updated ${packagePath}: ${newVersion}`);
}

function updateMainTs(newVersion) {
  const filePath = 'src/main.ts';
  let content = fs.readFileSync(filePath, 'utf8');
  
  // Оновлюємо версію в loadCurrentVersion методі
  content = content.replace(
    /if \(versionEl\) versionEl\.textContent = '[^']+'/,
    `if (versionEl) versionEl.textContent = '${newVersion}'`
  );
  
  fs.writeFileSync(filePath, content);
  console.log(`✅ Updated ${filePath}: ${newVersion}`);
}

function updateUpdateServiceTs(newVersion) {
  const filePath = 'electron/services/updateService.ts';
  let content = fs.readFileSync(filePath, 'utf8');
  
  // Оновлюємо версію в конструкторі UpdateService
  content = content.replace(
    /this\.currentVersion = '[^']+'/,
    `this.currentVersion = '${newVersion}'`
  );
  
  fs.writeFileSync(filePath, content);
  console.log(`✅ Updated ${filePath}: ${newVersion}`);
}

function updateReadmeAndDocs(oldVersion, newVersion) {
  // Оновлюємо файли документації
  const docsFiles = [
    'README.md',
    'RELEASE-NOTES-v1.1.0.md',
    'КОРИСТУВАЧАМ-ЗАВАНТАЖЕННЯ.md'
  ];

  docsFiles.forEach(filePath => {
    if (fs.existsSync(filePath)) {
      let content = fs.readFileSync(filePath, 'utf8');
      // Замінюємо всі згадки старої версії на нову
      content = content.replace(new RegExp(oldVersion.replace(/\./g, '\\.'), 'g'), newVersion);
      fs.writeFileSync(filePath, content);
      console.log(`✅ Updated ${filePath}: ${oldVersion} -> ${newVersion}`);
    }
  });
}

function main() {
  const currentVersion = getCurrentVersion();
  const newVersion = incrementPatchVersion(currentVersion);
  
  console.log(`🔄 Updating version: ${currentVersion} -> ${newVersion}`);
  
  // Оновлюємо всі файли
  updatePackageJson(newVersion);
  updateMainTs(newVersion);
  updateUpdateServiceTs(newVersion);
  updateReadmeAndDocs(currentVersion, newVersion);
  
  console.log(`\n🎉 Version successfully updated to ${newVersion}`);
  console.log(`📝 Next steps:`);
  console.log(`   1. npm run build && npm run package`);
  console.log(`   2. git add . && git commit -m "🔄 Auto-bump version to ${newVersion}"`);
  console.log(`   3. git tag v${newVersion} && git push origin v${newVersion}`);
}

if (require.main === module) {
  main();
}

module.exports = { getCurrentVersion, incrementPatchVersion };
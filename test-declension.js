// Простий тест для перевірки логіки відмінювання
const fs = require('fs');

// Імітуємо UkrainianNameDeclension клас
class TestUkrainianNameDeclension {
  
  // База чоловічих імен з відмінюванням
  static maleFirstNames = {
    "Олександр": {
      nominative: "Олександр", genitive: "Олександра", dative: "Олександру",
      accusative: "Олександра", instrumental: "Олександром", locative: "Олександрі"
    },
    "Володимир": {
      nominative: "Володимир", genitive: "Володимира", dative: "Володимиру",
      accusative: "Володимира", instrumental: "Володимиром", locative: "Володимирі"
    }
  };

  // База жіночих імен
  static femaleFirstNames = {
    "Олександра": {
      nominative: "Олександра", genitive: "Олександри", dative: "Олександрі",
      accusative: "Олександру", instrumental: "Олександрою", locative: "Олександрі"
    }
  };

  // База чоловічих по-батькові
  static malePatronymics = {
    "Володимирович": {
      nominative: "Володимирович", genitive: "Володимировича", dative: "Володимировичу",
      accusative: "Володимировича", instrumental: "Володимировичем", locative: "Володимировичі"
    }
  };

  // База прізвищ (універсальні закінчення)
  static surnamePatterns = {
    // Прізвища на -ак (наприклад Шостак)
    "ак": {
      male: (base) => ({
        nominative: base + "ак", genitive: base + "ака", dative: base + "аку",
        accusative: base + "ака", instrumental: base + "аком", locative: base + "аці"
      }),
      female: (base) => ({
        nominative: base + "ак", genitive: base + "ак", dative: base + "ак",
        accusative: base + "ак", instrumental: base + "ак", locative: base + "ак"
      })
    }
  };

  static getFirstNameForms(firstName) {
    const forms = new Set();
    forms.add(firstName);

    // Чоловічі імена
    if (this.maleFirstNames[firstName]) {
      const nameData = this.maleFirstNames[firstName];
      Object.values(nameData).forEach(form => forms.add(form));
    }

    // Жіночі імена  
    if (this.femaleFirstNames[firstName]) {
      const nameData = this.femaleFirstNames[firstName];
      Object.values(nameData).forEach(form => forms.add(form));
    }

    return Array.from(forms);
  }

  static getLastNameForms(lastName) {
    const forms = new Set();
    forms.add(lastName);

    // Перевіряємо шаблони прізвищ
    for (const [pattern, rules] of Object.entries(this.surnamePatterns)) {
      if (lastName.endsWith(pattern)) {
        const base = lastName.slice(0, -pattern.length);
        
        // Додаємо чоловічі форми
        const maleForms = rules.male(base);
        Object.values(maleForms).forEach(form => forms.add(form));
        
        // Додаємо жіночі форми
        const femaleForms = rules.female(base);
        Object.values(femaleForms).forEach(form => forms.add(form));
        
        break;
      }
    }

    return Array.from(forms);
  }

  static getMiddleNameForms(middleName) {
    const forms = new Set();
    forms.add(middleName);

    if (this.malePatronymics[middleName]) {
      const nameData = this.malePatronymics[middleName];
      Object.values(nameData).forEach(form => forms.add(form));
    }

    return Array.from(forms);
  }

  static getAllFormsOfName(fullName) {
    const parts = fullName.trim().split(/\s+/);
    if (parts.length < 2) return [fullName];

    // ВИПРАВЛЕНО: Для українських ПІБ порядок: Прізвище Ім'я По-батькові
    let lastName, firstName, middleName = '';
    
    if (parts.length === 2) {
      // Прізвище Ім'я
      [lastName, firstName] = parts;
    } else if (parts.length >= 3) {
      // Прізвище Ім'я По-батькові
      [lastName, firstName, middleName] = parts;
    } else {
      return [fullName];
    }

    const allForms = new Set();
    allForms.add(fullName); // Оригінальна форма

    console.log(`\n🔍 Розбираємо ПІБ: "${fullName}"`);
    console.log(`📝 Прізвище: "${lastName}", Ім'я: "${firstName}", По-батькові: "${middleName}"`);

    // Додаємо форми кожної частини імені
    const firstNameForms = this.getFirstNameForms(firstName);
    const lastNameForms = this.getLastNameForms(lastName);
    const middleNameForms = middleName ? this.getMiddleNameForms(middleName) : [''];

    console.log(`👤 Форми імені "${firstName}":`, firstNameForms);
    console.log(`🏷️ Форми прізвища "${lastName}":`, lastNameForms);
    console.log(`👨‍👦 Форми по-батькові "${middleName}":`, middleNameForms);

    // Генеруємо всі комбінації у правильному порядку: Прізвище Ім'я По-батькові
    lastNameForms.forEach(last => {
      firstNameForms.forEach(first => {
        if (middleName) {
          middleNameForms.forEach(middle => {
            allForms.add(`${last} ${first} ${middle}`);
            allForms.add(`${last} ${first}`); // Без по-батькові
            // Додаємо також варіанти з іншим порядком для більшої гнучкості
            allForms.add(`${first} ${middle} ${last}`);
            allForms.add(`${first} ${last}`);
          });
        } else {
          allForms.add(`${last} ${first}`);
          allForms.add(`${first} ${last}`); // Альтернативний порядок
        }
      });
    });

    const result = Array.from(allForms);
    console.log(`✅ Всі згенеровані форми (${result.length}):`, result);
    return result;
  }

  static findNameMatch(text, targetName) {
    const targetForms = this.getAllFormsOfName(targetName);
    const textLower = text.toLowerCase();
    
    console.log(`\n🔎 Шукаємо в тексті: "${text}"`);
    console.log(`🎯 Цільове ім'я: "${targetName}"`);
    
    // Спробувати знайти повне співпадіння спочатку
    for (const form of targetForms) {
      const words = form.split(' ').filter(w => w.length >= 3);
      console.log(`🔍 Перевіряємо форму: "${form}" -> слова:`, words);
      
      // Якщо всі слова знайдені - це ідеальний збіг
      if (words.length > 0 && words.every(word => {
        const found = textLower.includes(word.toLowerCase());
        console.log(`   - "${word}" ${found ? '✅ знайдено' : '❌ не знайдено'}`);
        return found;
      })) {
        console.log(`🎉 ЗБІГ! Форма "${form}" повністю знайдена в тексті`);
        return true;
      }
      
      // Якщо знайдено хоча б 2 слова з 3+ або 1 слово якщо воно довше 4 символів
      const foundWords = words.filter(word => textLower.includes(word.toLowerCase()));
      if (foundWords.length >= 2 || (foundWords.length >= 1 && foundWords[0].length > 4)) {
        console.log(`🎯 ЧАСТКОВИЙ ЗБІГ! Знайдено слова:`, foundWords);
        return true;
      }
    }
    
    console.log(`❌ Збігів не знайдено`);
    return false;
  }

  static findOrderParagraphs(wordText, excelNames) {
    console.log(`\n📊 === АНАЛІЗ WORD ДОКУМЕНТУ ===`);
    console.log(`📝 Текст для аналізу: "${wordText.substring(0, 200)}${wordText.length > 200 ? '...' : ''}"`);
    console.log(`📋 ПІБ з Excel (${excelNames.length}):`, excelNames);

    const paragraphs = wordText.split(/\n\s*\n/).filter(p => p.trim());
    console.log(`📄 Всього абзаців: ${paragraphs.length}`);

    const results = [];
    let currentPosition = 0;

    for (let i = 0; i < paragraphs.length; i++) {
      const paragraph = paragraphs[i];
      console.log(`\n🔍 === АБЗАЦ ${i + 1} ===`);
      console.log(`📝 Текст абзацу: "${paragraph}"`);
      
      const containsOrderKeyword = /розпорядженн[іїя]/i.test(paragraph);
      console.log(`🔑 Містить "розпорядженні": ${containsOrderKeyword ? '✅ ТАК' : '❌ НІ'}`);
      
      const matchedNames = [];

      if (containsOrderKeyword) {
        console.log(`🔍 Шукаємо ПІБ в цьому абзаці...`);
        
        // Шукаємо ПІБ з Excel в цьому абзаці
        for (const excelName of excelNames) {
          if (!excelName || excelName.trim() === '') continue;
          
          console.log(`\n👤 Перевіряємо ПІБ: "${excelName}"`);
          const found = this.findNameMatch(paragraph, excelName.trim());
          
          if (found) {
            matchedNames.push(excelName.trim());
            console.log(`✅ ПІБ "${excelName}" ЗНАЙДЕНО в абзаці!`);
          } else {
            console.log(`❌ ПІБ "${excelName}" НЕ знайдено в абзаці`);
          }
        }
      } else {
        console.log(`⏭️ Абзац не містить "розпорядженні", пропускаємо пошук ПІБ`);
      }

      const finallyIncluded = containsOrderKeyword && matchedNames.length > 0;
      console.log(`📊 РЕЗУЛЬТАТ АБЗАЦУ ${i + 1}:`);
      console.log(`   - Містить "розпорядженні": ${containsOrderKeyword}`);
      console.log(`   - Знайдені ПІБ (${matchedNames.length}):`, matchedNames);
      console.log(`   - ВКЛЮЧИТИ В РЕЗУЛЬТАТ: ${finallyIncluded ? '✅ ТАК' : '❌ НІ'}`);
      
      if (finallyIncluded) {
        results.push({
          paragraph: paragraph.trim(),
          containsOrderKeyword: true,
          matchedNames: [...new Set(matchedNames)],
          startPosition: currentPosition
        });
      }

      currentPosition += paragraph.length + 2; // +2 для \n\n
    }

    console.log(`\n🎯 === ПІДСУМОК ===`);
    console.log(`📊 Знайдено абзаців для результату: ${results.length}`);
    results.forEach((result, index) => {
      console.log(`   ${index + 1}. Знайдені ПІБ: ${result.matchedNames.join(', ')}`);
    });

    return results;
  }
}

// ТЕСТОВІ ДАНІ
console.log('🧪 === ТЕСТ ЛОГІКИ ВІДМІНЮВАННЯ ===\n');

// Тестові ПІБ з Excel (колонка D)
const excelNames = [
  "Шостак Олександр Володимирович",
  "Іваненко Володимир Петрович"
];

// Тестовий текст Word документу
const wordText = `
Першый абзац без ключових слів.

У цьому розпорядженні згадується Шостака Олександра Володимировича як відповідального за виконання.

Інший абзац також про розпорядження, але тут немає жодних імен з нашого списку.

Четвертий абзац взагалі не про розпорядження, хоча тут є Шостак Олександр.

П'ятий абзац містить розпорядження про Іваненка Володимира Петровича.
`;

console.log('📋 ТЕСТОВІ ДАНІ:');
console.log('Excel ПІБ:', excelNames);
console.log('Word текст:', wordText);

// Запускаємо тест
const results = TestUkrainianNameDeclension.findOrderParagraphs(wordText, excelNames);

console.log('\n🎯 === ФІНАЛЬНИЙ РЕЗУЛЬТАТ ===');
console.log(`Знайдено ${results.length} абзаців:`);
results.forEach((result, index) => {
  console.log(`\n${index + 1}. АБЗАЦ:`);
  console.log(`   Текст: "${result.paragraph}"`);
  console.log(`   Знайдені ПІБ: ${result.matchedNames.join(', ')}`);
});
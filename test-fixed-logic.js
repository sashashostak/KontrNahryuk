// Тест виправленої логіки пошуку ПІБ
const fs = require('fs');

// Копіюємо виправлену логіку з UkrainianNameDeclension
class FixedUkrainianNameDeclension {
  
  // База чоловічих імен
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

  // База чоловічих по-батькові
  static malePatronymics = {
    "Володимирович": {
      nominative: "Володимирович", genitive: "Володимировича", dative: "Володимировичу",
      accusative: "Володимировича", instrumental: "Володимировичем", locative: "Володимировичі"
    },
    "Петрович": {
      nominative: "Петрович", genitive: "Петровича", dative: "Петровичу", 
      accusative: "Петровича", instrumental: "Петровичем", locative: "Петровичі"
    }
  };

  // База прізвищ
  static surnamePatterns = {
    "ак": { // для Шостак
      male: (base) => ({
        nominative: base + "ак", genitive: base + "ака", dative: base + "аку",
        accusative: base + "ака", instrumental: base + "аком", locative: base + "аці"
      })
    },
    "енко": { // для Іваненко  
      male: (base) => ({
        nominative: base + "енко", genitive: base + "енко", dative: base + "енко",
        accusative: base + "енко", instrumental: base + "енко", locative: base + "енко"
      })
    }
  };

  static getFirstNameForms(firstName) {
    if (this.maleFirstNames[firstName]) {
      return Object.values(this.maleFirstNames[firstName]);
    }
    return [firstName]; // Якщо немає в базі, повертаємо як є
  }

  static getLastNameForms(lastName) {
    // Перевіряємо паттерни прізвищ
    for (const [pattern, rules] of Object.entries(this.surnamePatterns)) {
      if (lastName.endsWith(pattern)) {
        const base = lastName.slice(0, -pattern.length);
        if (rules.male) {
          return Object.values(rules.male(base));
        }
      }
    }
    return [lastName]; // Якщо не знайдено паттерн
  }

  static getMiddleNameForms(middleName) {
    if (this.malePatronymics[middleName]) {
      return Object.values(this.malePatronymics[middleName]);
    }
    return [middleName]; // Якщо немає в базі
  }

  static getAllFormsOfName(fullName) {
    const parts = fullName.trim().split(/\s+/);
    let lastName = '', firstName = '', middleName = '';

    if (parts.length >= 3) {
      [lastName, firstName, middleName] = parts;
    } else if (parts.length === 2) {
      [lastName, firstName] = parts;
    } else {
      return [fullName];
    }

    const lastNameForms = this.getLastNameForms(lastName);
    const firstNameForms = this.getFirstNameForms(firstName);
    const middleNameForms = middleName ? this.getMiddleNameForms(middleName) : [''];

    const allForms = [];

    // Генеруємо всі комбінації
    for (const lForm of lastNameForms) {
      for (const fForm of firstNameForms) {
        for (const mForm of middleNameForms) {
          if (middleName) {
            // З по-батькові
            allForms.push(`${lForm} ${fForm} ${mForm}`);
            allForms.push(`${fForm} ${mForm} ${lForm}`);
          }
          // Без по-батькові
          allForms.push(`${lForm} ${fForm}`);
          allForms.push(`${fForm} ${lForm}`);
        }
      }
    }

    return [...new Set(allForms)]; // Унікальні значення
  }

  // ВИПРАВЛЕНА логіка пошуку - потрібно знайти ВСІ слова
  static findNameMatch(text, targetName) {
    const targetForms = this.getAllFormsOfName(targetName);
    const textLower = text.toLowerCase();
    
    console.log(`\n🔎 Шукаємо в тексті: "${text}"`);
    console.log(`🎯 Цільове ім'я: "${targetName}"`);
    
    // Спробувати знайти точний збіг для кожної форми
    for (const form of targetForms) {
      const words = form.split(' ').filter(w => w.length >= 2);
      
      console.log(`🔍 Перевіряємо форму: "${form}" -> слова: [`, words.map(w => `'${w}'`).join(', '), `]`);
      
      // Перевіряємо кожне слово
      const foundWords = [];
      for (const word of words) {
        if (textLower.includes(word.toLowerCase())) {
          console.log(`   - "${word}" ✅ знайдено`);
          foundWords.push(word);
        } else {
          console.log(`   - "${word}" ❌ не знайдено`);
        }
      }
      
      // Якщо всі слова форми знайдені в тексті - це точний збіг
      if (words.length > 0 && foundWords.length === words.length) {
        console.log(`✅ ПІБ "${targetName}" ЗНАЙДЕНО в абзаці!`);
        return true;
      }
    }
    
    console.log(`❌ ПІБ "${targetName}" НЕ знайдено в абзаці`);
    return false;
  }

  static findOrderParagraphs(documentText, names) {
    console.log(`\n🔍 === ПОШУК РОЗПОРЯДЖЕНЬ ===`);
    console.log(`📋 Шукаємо ПІБ: [${names.map(n => `"${n}"`).join(', ')}]`);
    
    const paragraphs = documentText.split(/\n+/).filter(p => p.trim());
    const results = [];
    
    for (let i = 0; i < paragraphs.length; i++) {
      const paragraph = paragraphs[i];
      
      console.log(`\n🔍 === АБЗАЦ ${i + 1} ===`);
      console.log(`📝 Текст абзацу: "${paragraph}"`);
      
      // Перевіряємо наявність слова "розпорядження"
      const containsOrder = paragraph.toLowerCase().includes('розпорядження');
      console.log(`🔑 Містить "розпорядженні": ${containsOrder ? '✅ ТАК' : '❌ НІ'}`);
      
      if (containsOrder) {
        console.log(`🔍 Шукаємо ПІБ в цьому абзаці...`);
        
        const foundNames = [];
        for (const name of names) {
          console.log(`\n👤 Перевіряємо ПІБ: "${name}"`);
          
          if (this.findNameMatch(paragraph, name)) {
            foundNames.push(name);
          }
        }
        
        console.log(`📊 РЕЗУЛЬТАТ АБЗАЦУ ${i + 1}:`);
        console.log(`   - Містить "розпорядженні": ${containsOrder}`);
        console.log(`   - Знайдені ПІБ (${foundNames.length}): [`, foundNames.map(n => `'${n}'`).join(', '), `]`);
        
        if (foundNames.length > 0) {
          console.log(`   - ВКЛЮЧИТИ В РЕЗУЛЬТАТ: ✅ ТАК`);
          results.push({
            text: paragraph,
            foundNames: foundNames
          });
        } else {
          console.log(`   - ВКЛЮЧИТИ В РЕЗУЛЬТАТ: ❌ НІ`);
        }
      }
    }
    
    return results;
  }
}

// ===== ТЕСТ =====
console.log(`🧪 === ТЕСТ ВИПРАВЛЕНОЇ ЛОГІКИ ===\n`);

const testDocument = `Перший абзац не містить нічого особливого.

У цьому розпорядженні згадується Шостака Олександра Володимировича як відповідального за виконання.

Третій абзац також без розпорядження.

Четвертий абзац взагалі не про розпорядження, хоча тут є Шостак Олександр.

П'ятий абзац містить розпорядження про Іваненка Володимира Петровича.
`;

const testNames = [
  'Шостак Олександр Володимирович',
  'Іваненко Володимир Петрович'
];

const results = FixedUkrainianNameDeclension.findOrderParagraphs(testDocument, testNames);

console.log(`\n🎯 === ПІДСУМОК ===`);
console.log(`📊 Знайдено абзаців для результату: ${results.length}`);
results.forEach((result, index) => {
  console.log(`   ${index + 1}. Знайдені ПІБ: ${result.foundNames.join(', ')}`);
});

console.log(`\n🎯 === ФІНАЛЬНИЙ РЕЗУЛЬТАТ ===`);
if (results.length > 0) {
  console.log(`Знайдено ${results.length} абзаців:\n`);
  results.forEach((result, index) => {
    console.log(`${index + 1}. АБЗАЦ:`);
    console.log(`   Текст: "${result.text}"`);
    console.log(`   Знайдені ПІБ: ${result.foundNames.join(', ')}\n`);
  });
} else {
  console.log(`❌ Жодного підходящого абзацу не знайдено.`);
}
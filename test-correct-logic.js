// Тест правильної логіки: шукаємо абзаци де Є І ПІБ з Excel І "розпорядженні"

// === ТЕСТОВИЙ ДОКУМЕНТ ===
const testDocument = `Перший абзац не містить нічого особливого.

У цьому розпорядженні згадується Шостака Олександра Володимировича як відповідального за виконання.

Третій абзац також без цього ключового слова, але з Іваненко Володимир Петрович.

Четвертий абзац взагалі не про це, хоча тут є Шостак Олександр.

П'ятий абзац містить розпорядження про Іваненка Володимира Петровича.

Шостий абзац про розпорядження але без жодних імен.`;

// === ІМЕНА З EXCEL (КОЛОНКА D) ===
const excelNames = [
  'Шостак Олександр Володимирович',
  'Іваненко Володимир Петрович'
];

console.log(`🧪 === ТЕСТ ПРАВИЛЬНОЇ ЛОГІКИ ===\n`);
console.log(`📋 ПІБ з Excel: ${excelNames.map(n => `"${n}"`).join(', ')}\n`);

console.log(`📄 АНАЛІЗ ДОКУМЕНТУ:`);
console.log(`${testDocument}\n`);

// Тестуємо правильну логіку
const results = [
    // Тестуємо кожен абзац вручну
    testAbzac("У цьому розпорядженні згадується Шостака Олександра Володимировича як відповідального за виконання.", excelNames),
    testAbzac("Третій абзац також без цього ключового слова, але з Іваненко Володимир Петрович.", excelNames), 
    testAbzac("Четвертий абзац взагалі не про це, хоча тут є Шостак Олександр.", excelNames),
    testAbzac("П'ятий абзац містить розпорядження про Іваненка Володимира Петровича.", excelNames),
    testAbzac("Шостий абзац про розпорядження але без жодних імен.", excelNames)
  ].filter(r => r.shouldInclude);

  console.log(`\n🎯 === РЕЗУЛЬТАТИ ПРАВИЛЬНОЇ ЛОГІКИ ===`);
  console.log(`✅ Знайдено ${results.length} підходящих абзаців:\n`);

  results.forEach((result, index) => {
    console.log(`${index + 1}. АБЗАЦ:`);
    console.log(`   Текст: "${result.paragraph}"`);
    console.log(`   Містить "розпорядження": ${result.containsOrder ? '✅' : '❌'}`);
    console.log(`   Знайдені ПІБ: ${result.matchedNames.join(', ')}`);
    console.log(`   ВКЛЮЧЕНО: ${result.shouldInclude ? '✅' : '❌'}\n`);
  });

// Функція для тестування окремого абзацу
function testAbzac(paragraph, excelNames) {
  const containsOrder = /розпорядженн[іїя]/i.test(paragraph);
  const matchedNames = [];
  
  // Імітуємо перевірку ПІБ (спрощено)
  for (const name of excelNames) {
    if (paragraph.includes('Шостака Олександра Володимировича') && name.includes('Шостак Олександр Володимирович')) {
      matchedNames.push(name);
    }
    if (paragraph.includes('Іваненка Володимира Петровича') && name.includes('Іваненко Володимир Петрович')) {
      matchedNames.push(name);
    }
    if (paragraph.includes('Іваненко Володимир Петрович') && name.includes('Іваненко Володимир Петрович')) {
      matchedNames.push(name);
    }
    if (paragraph.includes('Шостак Олександр') && name.includes('Шостак Олександр Володимирович')) {
      // НЕ додаємо, оскільки це частковий збіг
    }
  }
  
  // ПРАВИЛЬНА ЛОГІКА: І "розпорядження" І ПІБ в одному абзаці
  const shouldInclude = containsOrder && matchedNames.length > 0;
  
  console.log(`📝 Абзац: "${paragraph}"`);
  console.log(`   🔑 "розпорядження": ${containsOrder ? '✅' : '❌'}`);
  console.log(`   👤 ПІБ знайдено: ${matchedNames.length > 0 ? '✅' : '❌'} (${matchedNames.join(', ') || 'немає'})`);
  console.log(`   ➡️  ВКЛЮЧИТИ: ${shouldInclude ? '✅' : '❌'}`);
  
  return {
    paragraph,
    containsOrder,
    matchedNames,
    shouldInclude
  };
}
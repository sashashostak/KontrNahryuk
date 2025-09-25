const fs = require('fs')
const path = require('path')
const mammoth = require('mammoth')
const { Document, Paragraph, Packer, TextRun } = require('docx')

async function testFormatting() {
  try {
    // Читаємо test-word.docx
    const inputPath = path.join(__dirname, 'test-word.docx')
    const outputPath = path.join(__dirname, 'test-output.docx')
    
    console.log('Читаємо документ...')
    
    const wordBuf = fs.readFileSync(inputPath)
    const result = await mammoth.extractRawText(wordBuf)
    const paragraphs = result.value.split('\n').filter(p => p.trim())
    
    console.log(`Знайдено ${paragraphs.length} абзаців`)
    
    // Пошук за ключовим словом "токен1_1"
    const keyword = "токен1_1"
    const matchedParagraphs = []
    
    for (let i = 0; i < paragraphs.length; i++) {
      const text = paragraphs[i]
      if (text.toLowerCase().includes(keyword.toLowerCase())) {
        // Додаємо попередній абзац (контекст)
        if (i > 0) {
          matchedParagraphs.push({
            text: paragraphs[i-1],
            type: 'context'
          })
        }
        
        // Додаємо знайдений абзац
        matchedParagraphs.push({
          text: text,
          type: 'match'
        })
        
        // Додаємо наступний абзац (контекст)
        if (i < paragraphs.length - 1) {
          matchedParagraphs.push({
            text: paragraphs[i+1],
            type: 'context'
          })
        }
      }
    }
    
    console.log(`Знайдено ${matchedParagraphs.length} абзаців з контекстом`)
    
    // Створюємо Word документ
    const docChildren = []
    
    // Додаємо заголовок
    docChildren.push(new Paragraph({
      children: [new TextRun({
        text: "Результати пошуку за ключовим словом: " + keyword,
        font: "Calibri",
        size: 28,
        bold: true
      })],
      alignment: 'both', // Вирівнювання за шириною
      spacing: { after: 400 }
    }))
    
    // Додаємо знайдені абзаци
    for (let i = 0; i < matchedParagraphs.length; i++) {
      const item = matchedParagraphs[i]
      const nextItem = i < matchedParagraphs.length - 1 ? matchedParagraphs[i + 1] : null
      
      // Визначаємо тип абзацу
      const text = item.text.trim()
      let isBold = false
      let isPointOrSubpoint = false
      let isPoint = false
      
      // Перевіряємо чи це пункт (1. 2. 3.)
      if (text.match(/^\d+\.\s+/)) {
        isBold = true
        isPointOrSubpoint = true
        isPoint = true
      }
      // Перевіряємо чи це підпункт (1.1. 2.3.)
      else if (text.match(/^\d+\.\d+\.?\s+/)) {
        isBold = true
        isPointOrSubpoint = true
        isPoint = false
      }
      // Інше - звичайний абзац
      else {
        isBold = false
      }
      
      // Підсвічуємо знайдені збіги
      if (item.type === 'match') {
        isBold = true
      }
      
      // Додати пустий рядок перед пунктами та підпунктами
      if (isPointOrSubpoint) {
        docChildren.push(new Paragraph({
          children: [new TextRun({ text: "", font: "Calibri", size: 28 })],
          alignment: 'both' // Вирівнювання за шириною
        }))
      }
      
      // Додати основний абзац
      docChildren.push(new Paragraph({
        children: [new TextRun({
          text: text,
          font: "Calibri",
          size: 28, // 14pt
          bold: isBold
        })],
        alignment: 'both', // Вирівнювання за шириною
        indent: {
          firstLine: 720 // Абзацний відступ (0.5 дюйма)
        }
        // Ніяких spacing
      }))
      
      // Додати пустий рядок після пунктів та підпунктів,
      // але НЕ додавати, якщо наступний елемент - підпункт
      if (isPointOrSubpoint) {
        const nextText = nextItem ? nextItem.text.trim() : ""
        const nextIsSubpoint = nextText.match(/^\d+\.\d+\.?\s+/)
        const shouldAddEmptyLine = !(isPoint && nextIsSubpoint)
        
        if (shouldAddEmptyLine) {
          docChildren.push(new Paragraph({
            children: [new TextRun({ text: "", font: "Calibri", size: 28 })],
            alignment: 'both' // Вирівнювання за шириною
          }))
        }
      }
    }
    
    // Створюємо та зберігаємо документ
    const doc = new Document({
      sections: [{
        properties: {},
        children: docChildren
      }]
    })
    
    const buffer = await Packer.toBuffer(doc)
    fs.writeFileSync(outputPath, buffer)
    
    console.log(`Документ збережено: ${outputPath}`)
    console.log(`Знайдено збігів: ${matchedParagraphs.filter(p => p.type === 'match').length}`)
    
  } catch (error) {
    console.error('Помилка:', error)
  }
}

testFormatting()
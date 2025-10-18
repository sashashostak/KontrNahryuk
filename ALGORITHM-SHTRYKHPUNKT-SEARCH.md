# 🔍 АЛГОРИТМ ПОШУКУ ШТРИХПУНКТІВ

## Поточна логіка (що реалізовано в коді)

### Крок 1: Конвертація Word → HTML
```typescript
const result = await mammoth.convertToHtml({ 
  buffer: Buffer.from(wordBuf)
})
```
**Результат:** Весь документ стає одним HTML рядком

---

### Крок 2: Розбиття на абзаци
```typescript
const htmlParagraphs = result.value.split(/<\/?p[^>]*>/i).filter(p => p.trim().length > 0)
```
**Що робить:** Розділяє по тегах `<p>` і `</p>`, видаляє порожні

**Приклад:**
```html
<p>Наказ №123</p><p><u>молодший сержант</u></p><p>солдата ІВАНОВА...</p>
```
**Стає:**
```
[0] "Наказ №123"
[1] "<u>молодший сержант</u>"
[2] "солдата ІВАНОВА..."
```

---

### Крок 3: Перевірка КОЖНОГО абзацу

Для кожного абзацу код перевіряє:

#### 3.1. ЧИ Є ПІДКРЕСЛЕННЯ? (6 варіантів)

```typescript
// Варіант 1: Тег <u>
const hasUnderlineTag = /<u[\s/>]/i.test(html)  // <u>, <u >, <u/>

// Варіант 2: Тег <u з атрибутами
const hasUnderlineTag2 = /<u\s+/i.test(html)    // <u class="...">

// Варіант 3: CSS text-decoration
const hasUnderlineStyle1 = /text-decoration\s*:\s*underline/i.test(html)

// Варіант 4: CSS text-decoration-line
const hasUnderlineStyle2 = /text-decoration-line\s*:\s*underline/i.test(html)

// Варіант 5: Inline style з подвійними лапками
const hasUnderlineStyle3 = /style="[^"]*underline/i.test(html)

// Варіант 6: Inline style з одинарними лапками
const hasUnderlineStyle4 = /style='[^']*underline/i.test(html)

// ЗАГАЛЬНИЙ РЕЗУЛЬТАТ:
const hasUnderline = hasUnderlineTag || hasUnderlineTag2 || 
                    hasUnderlineStyle1 || hasUnderlineStyle2 || 
                    hasUnderlineStyle3 || hasUnderlineStyle4
```

**Приклади HTML що мають спрацювати:**
```html
<u>молодший сержант</u>
<u class="underline">молодший сержант</u>
<span style="text-decoration: underline">молодший сержант</span>
<span style="text-decoration-line: underline">молодший сержант</span>
<span style="text-decoration: underline; color: red">молодший сержант</span>
<span style='text-decoration: underline'>молодший сержант</span>
```

#### 3.2. ЧИ Є ЖИРНИЙ? (6 варіантів)

```typescript
// Варіант 1: Теги <strong> або <b>
const hasBoldTag1 = /<(strong|b)[\s/>]/i.test(html)

// Варіант 2: Теги <strong> або <b> з атрибутами
const hasBoldTag2 = /<(strong|b)\s+/i.test(html)

// Варіант 3: CSS font-weight: bold
const hasBoldStyle1 = /font-weight\s*:\s*bold/i.test(html)

// Варіант 4: CSS font-weight: 600-900
const hasBoldStyle2 = /font-weight\s*:\s*[6-9]00/i.test(html)

// Варіант 5: Inline style з "bold" (подвійні лапки)
const hasBoldStyle3 = /style="[^"]*bold/i.test(html)

// Варіант 6: Inline style з "bold" (одинарні лапки)
const hasBoldStyle4 = /style='[^']*bold/i.test(html)

// ЗАГАЛЬНИЙ РЕЗУЛЬТАТ:
const hasBold = hasBoldTag1 || hasBoldTag2 || 
               hasBoldStyle1 || hasBoldStyle2 || 
               hasBoldStyle3 || hasBoldStyle4
```

**Приклади HTML що мають спрацювати як ЖИРНИЙ:**
```html
<strong>текст</strong>
<b>текст</b>
<b class="bold">текст</b>
<span style="font-weight: bold">текст</span>
<span style="font-weight: 600">текст</span>
<span style="font-weight: 700">текст</span>
<span style='font-weight: bold'>текст</span>
```

#### 3.3. ВИРІШАЛЬНА ПЕРЕВІРКА

```typescript
// ШтрихПункт = ТІЛЬКИ підкреслений БЕЗ жирного
const isBoldAndUnderlined = hasUnderline && !hasBold
```

**Логіка:**
- ✅ `hasUnderline = true` + `hasBold = false` → **ШтрихПункт!**
- ❌ `hasUnderline = true` + `hasBold = true` → **НЕ ШтрихПункт** (підкреслений + жирний)
- ❌ `hasUnderline = false` → **НЕ ШтрихПункт** (не підкреслений)

---

## 📊 Що логується в консоль

### Випадок 1: Знайдено ШтрихПункт
```
✅✅✅ ЗНАЙДЕНО ШТРИХПУНКТ на позиції 52:
   Текст: "молодший сержант"
   HTML: <u>молодший сержант</u>
   Underline: true (tag: true, style: false)
   Bold: false (має бути FALSE!)
```

### Випадок 2: Підкреслений + Жирний (НЕ ШтрихПункт)
```
⚠️  Пропущено (підкреслений + жирний) на позиції 48:
   Текст: "старший сержант..."
   Underline: true, Bold: true ❌
```

### Випадок 3: Підкреслений без жирного, але не розпізнано
```
🤔 Потенційний ШтрихПункт (перевірте) на позиції 55:
   Текст: "капітан"
   HTML: <u>капітан</u>
   Underline: true, Bold: false
   Чому не розпізнано? (можливо, занадто довгий текст або інша причина)
```

### Випадок 4: Короткий абзац без форматування
```
📝 Короткий абзац 60 (НЕ ШтрихПункт):
   Текст: "майор"
   HTML: майор
   Underline: false, Bold: false
```

---

## 🎯 КРИТИЧНІ ПИТАННЯ ДЛЯ КОРИСТУВАЧА

### 1. Як виглядає ШтрихПункт у Word?

**Опишіть формат:**
- Чи це звичайний текст з підкресленням (Ctrl+U)?
- Чи це спеціальний символ або маркер?
- Чи він у окремому абзаці чи всередині тексту?

**Приклад:**
```
13. ОГОЛОСИТИ про присвоєння військових звань:
молодший сержант  ← ЦЕ ШТРИХПУНКТ?
головному сержанту ІВАНОВУ Іван Іванович, 2-го батальйону
старший сержант   ← ЦЕ ШТРИХПУНКТ?
сержанту ПЕТРОВУ Петро Петрович, 2-го батальйону
```

### 2. Чи є інші ознаки ШтрихПункту?

**Можливі додаткові критерії:**
- Коротка довжина тексту? (наприклад, < 50 символів)
- Конкретні слова? (звання: "сержант", "лейтенант", "капітан")
- Положення в структурі? (йде після пункту/підпункту)
- Регулярний вираз тексту? (наприклад, закінчується на певне слово)

### 3. Приклад HTML який НЕ працює

**Надайте фактичний HTML з консолі:**

Коли запустите додаток і обробите документ, в консолі буде:
```
[extractFormatted] === ПОВНИЙ СПИСОК HTML АБЗАЦІВ (перші 20) ===
[0] ...
[1] ...
[52] <ЯКИЙСЬ HTML>  ← СКОПІЮЙТЕ ЦЕЙ РЯДОК!
```

Якщо бачите ШтрихПункт але він НЕ розпізнається — **скопіюйте його HTML!**

---

## 🔧 МОЖЛИВІ ВИПРАВЛЕННЯ

### Якщо ШтрихПункт має специфічний формат:

#### Варіант A: Додати перевірку довжини
```typescript
const isBoldAndUnderlined = hasUnderline && !hasBold && text.length < 50
```

#### Варіант B: Перевірити за ключовими словами
```typescript
const militaryRanks = /сержант|лейтенант|капітан|майор|полковник/i
const isBoldAndUnderlined = hasUnderline && !hasBold && militaryRanks.test(text)
```

#### Варіант C: Перевірити за структурою
```typescript
// Якщо йде після пункту (13., 15.1.) і короткий
const isBoldAndUnderlined = hasUnderline && !hasBold && text.length < 100
```

#### Варіант D: Специфічний HTML тег/атрибут
```typescript
// Якщо mammoth генерує щось специфічне
const hasSpecialClass = /<u class="dash-point">/i.test(html)
const isBoldAndUnderlined = hasSpecialClass
```

---

## 📝 ЩО МЕН ПОТРІБНО ВІД ВАС

### 1️⃣ Опис формату ШтрихПункту
Розкажіть своїми словами:
- Як він виглядає в Word
- Як його створили (які кнопки натиснули)
- Чи є щось спільне у всіх ШтрихПунктах

### 2️⃣ Скріншот з Word
Зробіть скріншот абзацу з ШтрихПунктом у Word

### 3️⃣ HTML з консолі
Запустіть додаток, обробіть документ, скопіюйте з консолі:
```
[extractFormatted] === ПОВНИЙ СПИСОК HTML АБЗАЦІВ (перші 20) ===
[0] ...
[52] <ТУТ МАЄ БУТИ ШТРИХПУНКТ> ← СКОПІЮЙТЕ
```

### 4️⃣ Блок діагностики
Якщо є, скопіюйте:
```
⚠️  Пропущено (підкреслений + жирний)
```
або
```
🤔 Потенційний ШтрихПункт (перевірте)
```

---

## 🚀 ШВИДКИЙ ТЕСТ

### Запустіть:
```powershell
npm run build
npm run dev
```

### Відкрийте консоль: Ctrl+Shift+I

### Обробіть документ

### Шукайте один з цих варіантів:

✅ **Успіх:**
```
✅✅✅ ЗНАЙДЕНО ШТРИХПУНКТ
ПІДСУМОК: Знайдено ШтрихПунктів: 7
```

⚠️ **Проблема:**
```
⚠️  Пропущено (підкреслений + жирний)
ПІДСУМОК: Знайдено ШтрихПунктів: 0
```

🤔 **Потенційна проблема:**
```
🤔 Потенційний ШтрихПункт (перевірте)
```

---

**Скопіюйте все що бачите в консолі і надішліть мені!** 📋

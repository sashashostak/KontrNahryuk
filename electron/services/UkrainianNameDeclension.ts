/**
 * Бібліотека відмінювання українських імен, прізвищ та по-батькові
 * Для правильного зіставлення ПІБ у різних відмінках
 * 
 * ВАЖЛИВО: У режимі "Розпорядження" ПІБ шукаються в колонці D Excel файлу
 * - Excel колонка D містить повні ПІБ у називному відмінку
 * - Word документи містять ПІБ в різних відмінках (родовий, давальний тощо)
 * - Бібліотека дозволяє знайти відповідність між ними
 */

export interface NameDeclension {
  nominative: string;    // Іменний відмінок (Хто? Що?) - Іван Петренко
  genitive: string;      // Родовий відмінок (Кого? Чого?) - Івана Петренка  
  dative: string;        // Давальний відмінок (Кому? Чому?) - Івану Петренку
  accusative: string;    // Знахідний відмінок (Кого? Що?) - Івана Петренка
  locative: string;      // Місцевий відмінок (На кому? На чому?) - на Івані Петренку
}

export class UkrainianNameDeclension {
  
  // База чоловічих імен з відмінюванням
  private static maleFirstNames: Record<string, NameDeclension> = {
    "Олександр": {
      nominative: "Олександр", genitive: "Олександра", dative: "Олександру",
      accusative: "Олександра", locative: "Олександрі"
    },
    "Володимир": {
      nominative: "Володимир", genitive: "Володимира", dative: "Володимиру",
      accusative: "Володимира", locative: "Володимирі"
    },
    "Іван": {
      nominative: "Іван", genitive: "Івана", dative: "Івану",
      accusative: "Івана", locative: "Івані"
    },
    "Сергій": {
      nominative: "Сергій", genitive: "Сергія", dative: "Сергію",
      accusative: "Сергія", locative: "Сергії"
    },
    "Андрій": {
      nominative: "Андрій", genitive: "Андрія", dative: "Андрію",
      accusative: "Андрія", locative: "Андрії"
    },
    "Дмитро": {
      nominative: "Дмитро", genitive: "Дмитра", dative: "Дмитру",
      accusative: "Дмитра", locative: "Дмитрі"
    },
    "Віктор": {
      nominative: "Віктор", genitive: "Віктора", dative: "Віктору",
      accusative: "Віктора", locative: "Вікторі"
    },
    "Михайло": {
      nominative: "Михайло", genitive: "Михайла", dative: "Михайлу",
      accusative: "Михайла", locative: "Михайлі"
    },
    "Олег": {
      nominative: "Олег", genitive: "Олега", dative: "Олегу",
      accusative: "Олега", locative: "Олезі"
    },
    "Юрій": {
      nominative: "Юрій", genitive: "Юрія", dative: "Юрію",
      accusative: "Юрія", locative: "Юрії"
    },
    "Василь": {
      nominative: "Василь", genitive: "Василя", dative: "Василю",
      accusative: "Василя", locative: "Василі"
    },
    "Петро": {
      nominative: "Петро", genitive: "Петра", dative: "Петру",
      accusative: "Петра", locative: "Петрі"
    },
    "Артем": {
      nominative: "Артем", genitive: "Артема", dative: "Артему",
      accusative: "Артема", locative: "Артемі"
    },
    "Ігор": {
      nominative: "Ігор", genitive: "Ігоря", dative: "Ігорю",
      accusative: "Ігоря", locative: "Ігорі"
    },
    "Роман": {
      nominative: "Роман", genitive: "Романа", dative: "Роману",
      accusative: "Романа", locative: "Романі"
    },
    "Максим": {
      nominative: "Максим", genitive: "Максима", dative: "Максиму",
      accusative: "Максима", locative: "Максимі"
    }
  };

  // База жіночих імен з відмінюванням
  private static femaleFirstNames: Record<string, NameDeclension> = {
    "Олена": {
      nominative: "Олена", genitive: "Олени", dative: "Олені",
      accusative: "Олену", locative: "Олені"
    },
    "Тетяна": {
      nominative: "Тетяна", genitive: "Тетяни", dative: "Тетяні",
      accusative: "Тетяну", locative: "Тетяні"
    },
    "Наталія": {
      nominative: "Наталія", genitive: "Наталії", dative: "Наталії",
      accusative: "Наталію", locative: "Наталії"
    },
    "Ірина": {
      nominative: "Ірина", genitive: "Ірини", dative: "Ірині",
      accusative: "Ірину", locative: "Ірині"
    },
    "Світлана": {
      nominative: "Світлана", genitive: "Світлани", dative: "Світлані",
      accusative: "Світлану", locative: "Світлані"
    },
    "Людмила": {
      nominative: "Людмила", genitive: "Людмили", dative: "Людмилі",
      accusative: "Людмилу", locative: "Людмилі"
    },
    "Валентина": {
      nominative: "Валентина", genitive: "Валентини", dative: "Валентині",
      accusative: "Валентину", locative: "Валентині"
    },
    "Марія": {
      nominative: "Марія", genitive: "Марії", dative: "Марії",
      accusative: "Марію", locative: "Марії"
    },
    "Анна": {
      nominative: "Анна", genitive: "Анни", dative: "Анні",
      accusative: "Анну", locative: "Анні"
    },
    "Катерина": {
      nominative: "Катерина", genitive: "Катерини", dative: "Катерині",
      accusative: "Катерину", locative: "Катерині"
    }
  };

  // База чоловічих по-батькові
  private static malePatronymics: Record<string, NameDeclension> = {
    "Олександрович": {
      nominative: "Олександрович", genitive: "Олександровича", dative: "Олександровичу",
      accusative: "Олександровича", locative: "Олександровичі"
    },
    "Володимирович": {
      nominative: "Володимирович", genitive: "Володимировича", dative: "Володимировичу",
      accusative: "Володимировича", locative: "Володимировичі"
    },
    "Іванович": {
      nominative: "Іванович", genitive: "Івановича", dative: "Івановичу",
      accusative: "Івановича", locative: "Івановичі"
    },
    "Сергійович": {
      nominative: "Сергійович", genitive: "Сергійовича", dative: "Сергійовичу",
      accusative: "Сергійовича", locative: "Сергійовичі"
    },
    "Андрійович": {
      nominative: "Андрійович", genitive: "Андрійовича", dative: "Андрійовичу",
      accusative: "Андрійовича", locative: "Андрійовичі"
    },
    "Дмитрович": {
      nominative: "Дмитрович", genitive: "Дмитровича", dative: "Дмитровичу",
      accusative: "Дмитровича", locative: "Дмитровичі"
    },
    "Вікторович": {
      nominative: "Вікторович", genitive: "Вікторовича", dative: "Вікторовичу",
      accusative: "Вікторовича", locative: "Вікторовичі"
    },
    "Михайлович": {
      nominative: "Михайлович", genitive: "Михайловича", dative: "Михайловичу",
      accusative: "Михайловича", locative: "Михайловичі"
    }
  };

  // База жіночих по-батькові
  private static femalePatronymics: Record<string, NameDeclension> = {
    "Олександрівна": {
      nominative: "Олександрівна", genitive: "Олександрівни", dative: "Олександрівні",
      accusative: "Олександрівну", locative: "Олександрівні"
    },
    "Володимирівна": {
      nominative: "Володимирівна", genitive: "Володимирівни", dative: "Володимирівні",
      accusative: "Володимирівну", locative: "Володимирівні"
    },
    "Іванівна": {
      nominative: "Іванівна", genitive: "Іванівни", dative: "Іванівні",
      accusative: "Іванівну", locative: "Іванівні"
    },
    "Сергіївна": {
      nominative: "Сергіївна", genitive: "Сергіївни", dative: "Сергіївні",
      accusative: "Сергіївну", locative: "Сергіївні"
    },
    "Андріївна": {
      nominative: "Андріївна", genitive: "Андріївни", dative: "Андріївні",
      accusative: "Андріївну", locative: "Андріївні"
    },
    "Дмитрівна": {
      nominative: "Дмитрівна", genitive: "Дмитрівни", dative: "Дмитрівні",
      accusative: "Дмитрівну", locative: "Дмитрівні"
    },
    "Вікторівна": {
      nominative: "Вікторівна", genitive: "Вікторівни", dative: "Вікторівні",
      accusative: "Вікторівну", locative: "Вікторівні"
    },
    "Михайлівна": {
      nominative: "Михайлівна", genitive: "Михайлівни", dative: "Михайлівні",
      accusative: "Михайлівну", locative: "Михайлівні"
    }
  };

  // База прізвищ (універсальні закінчення)
  private static surnamePatterns = {
    // Чоловічі прізвища на -ко
    "ко": {
      male: (base: string) => ({
        nominative: base + "ко", genitive: base + "ка", dative: base + "ку",
        accusative: base + "ка", instrumental: base + "ком", locative: base + "ку"
      }),
      female: (base: string) => ({
        nominative: base + "ко", genitive: base + "ко", dative: base + "ко",
        accusative: base + "ко", instrumental: base + "ко", locative: base + "ко"
      })
    },
    // Прізвища на -енко/-enko
    "енко": {
      male: (base: string) => ({
        nominative: base + "енко", genitive: base + "енка", dative: base + "енку",
        accusative: base + "енка", instrumental: base + "енком", locative: base + "енку"
      }),
      female: (base: string) => ({
        nominative: base + "енко", genitive: base + "енко", dative: base + "енко",
        accusative: base + "енко", instrumental: base + "енко", locative: base + "енко"
      })
    },
    // Прізвища на -ський/-цький
    "ський": {
      male: (base: string) => ({
        nominative: base + "ський", genitive: base + "ського", dative: base + "ському",
        accusative: base + "ського", instrumental: base + "ським", locative: base + "ському"
      }),
      female: (base: string) => ({
        nominative: base + "ська", genitive: base + "ської", dative: base + "ській",
        accusative: base + "ську", instrumental: base + "ською", locative: base + "ській"
      })
    },
    // Прізвища на -ич/-ич
    "ич": {
      male: (base: string) => ({
        nominative: base + "ич", genitive: base + "ича", dative: base + "ичу",
        accusative: base + "ича", instrumental: base + "ичем", locative: base + "ичі"
      }),
      female: (base: string) => ({
        nominative: base + "ич", genitive: base + "ич", dative: base + "ич",
        accusative: base + "ич", instrumental: base + "ич", locative: base + "ич"
      })
    },
    // Прізвища на -ак (наприклад Шостак)
    "ак": {
      male: (base: string) => ({
        nominative: base + "ак", genitive: base + "ака", dative: base + "аку",
        accusative: base + "ака", instrumental: base + "аком", locative: base + "аці"
      }),
      female: (base: string) => ({
        nominative: base + "ак", genitive: base + "ак", dative: base + "ак",
        accusative: base + "ак", instrumental: base + "ак", locative: base + "ак"
      })
    },
    // Прізвища на -юк/-юк (наприклад Савчук)
    "юк": {
      male: (base: string) => ({
        nominative: base + "юк", genitive: base + "юка", dative: base + "юку",
        accusative: base + "юка", instrumental: base + "юком", locative: base + "юці"
      }),
      female: (base: string) => ({
        nominative: base + "юк", genitive: base + "юк", dative: base + "юк",
        accusative: base + "юк", instrumental: base + "юк", locative: base + "юк"
      })
    },
    // Прізвища на -ук (наприклад Тарасюк)
    "ук": {
      male: (base: string) => ({
        nominative: base + "ук", genitive: base + "ука", dative: base + "уку",
        accusative: base + "ука", instrumental: base + "уком", locative: base + "уці"
      }),
      female: (base: string) => ({
        nominative: base + "ук", genitive: base + "ук", dative: base + "ук",
        accusative: base + "ук", instrumental: base + "ук", locative: base + "ук"
      })
    },
    // Прізвища на -ський без бази (вже є "ський")
    "цький": {
      male: (base: string) => ({
        nominative: base + "цький", genitive: base + "цького", dative: base + "цькому",
        accusative: base + "цького", instrumental: base + "цьким", locative: base + "цькому"
      }),
      female: (base: string) => ({
        nominative: base + "цька", genitive: base + "цької", dative: base + "цькій",
        accusative: base + "цьку", instrumental: base + "цькою", locative: base + "цькій"
      })
    }
  };

  /**
   * Отримати всі можливі форми імені для пошуку
   */
  public static getAllFormsOfName(fullName: string): string[] {
    const parts = fullName.trim().split(/\s+/);
    if (parts.length < 2) return [fullName];

    // ВИПРАВЛЕНО: Для українських ПІБ порядок: Прізвище Ім'я По-батькові
    let lastName: string, firstName: string, middleName: string = '';
    
    if (parts.length === 2) {
      // Прізвище Ім'я
      [lastName, firstName] = parts;
    } else if (parts.length >= 3) {
      // Прізвище Ім'я По-батькові
      [lastName, firstName, middleName] = parts;
    } else {
      return [fullName];
    }

    const allForms = new Set<string>();
    allForms.add(fullName); // Оригінальна форма

    // Додаємо форми кожної частини імені
    const firstNameForms = this.getFirstNameForms(firstName);
    const lastNameForms = this.getLastNameForms(lastName);
    const middleNameForms = middleName ? this.getMiddleNameForms(middleName) : [''];

    // Генеруємо всі комбінації у правильному порядку: Прізвище Ім'я По-батькові
    lastNameForms.forEach(last => {
      firstNameForms.forEach(first => {
        if (middleName) {
          middleNameForms.forEach(middle => {
            allForms.add(`${last} ${first} ${middle}`);
            allForms.add(`${last} ${first}`); // Без по-батькові
          });
        } else {
          allForms.add(`${last} ${first}`);
        }
      });
    });

    return Array.from(allForms);
  }

  /**
   * Отримати всі форми імені
   */
  private static getFirstNameForms(name: string): string[] {
    // Спочатку шукаємо в базі чоловічих імен
    const maleForms = this.maleFirstNames[name];
    if (maleForms) {
      return Object.values(maleForms);
    }

    // Потім в базі жіночих імен
    const femaleForms = this.femaleFirstNames[name];
    if (femaleForms) {
      return Object.values(femaleForms);
    }

    // Якщо не знайдено в базі, застосовуємо базові правила
    return this.declineFirstNameByRules(name);
  }

  /**
   * Отримати всі форми по-батькові
   */
  private static getMiddleNameForms(middleName: string): string[] {
    // Чоловічі по-батькові
    const maleForms = this.malePatronymics[middleName];
    if (maleForms) {
      return Object.values(maleForms);
    }

    // Жіночі по-батькові
    const femaleForms = this.femalePatronymics[middleName];
    if (femaleForms) {
      return Object.values(femaleForms);
    }

    // Базові правила для по-батькові
    return this.declinePatronymicByRules(middleName);
  }

  /**
   * Отримати всі форми прізвища
   */
  private static getLastNameForms(lastName: string): string[] {
    const forms = new Set<string>();
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

  /**
   * Базові правила відмінювання імен
   */
  private static declineFirstNameByRules(name: string): string[] {
    const forms = [name]; // Початкова форма

    // Базові правила для чоловічих імен
    if (name.endsWith('о')) {
      // Дмитро -> Дмитра, Дмитру
      const base = name.slice(0, -1);
      forms.push(base + 'а', base + 'у', base + 'ом', base + 'і');
    } else if (name.endsWith('й')) {
      // Сергій -> Сергія, Сергію
      const base = name.slice(0, -1);
      forms.push(base + 'я', base + 'ю', base + 'єм', base + 'ї');
    } else if (!name.endsWith('а') && !name.endsWith('я')) {
      // Іван -> Івана, Івану
      forms.push(name + 'а', name + 'у', name + 'ом', name + 'і');
    }

    // Базові правила для жіночих імен
    if (name.endsWith('а')) {
      // Олена -> Олени, Олені
      const base = name.slice(0, -1);
      forms.push(base + 'и', base + 'і', base + 'у', base + 'ою');
    } else if (name.endsWith('я')) {
      // Марія -> Марії
      const base = name.slice(0, -1);
      forms.push(base + 'ї', base + 'ю', base + 'єю');
    }

    return forms;
  }

  /**
   * Базові правила відмінювання по-батькові
   */
  private static declinePatronymicByRules(patronymic: string): string[] {
    const forms = [patronymic];

    if (patronymic.endsWith('ович')) {
      const base = patronymic.slice(0, -4);
      forms.push(
        base + 'овича', base + 'овичу', base + 'овичем', base + 'овичі'
      );
    } else if (patronymic.endsWith('івна')) {
      const base = patronymic.slice(0, -4);
      forms.push(
        base + 'івни', base + 'івні', base + 'івну', base + 'івною'
      );
    }

    return forms;
  }

  /**
   * Перевірити, чи збігаються імена в різних відмінках
   */
  public static namesMatch(name1: string, name2: string): boolean {
    if (!name1 || !name2) return false;
    
    // Пряме співставлення
    if (name1.toLowerCase() === name2.toLowerCase()) return true;

    // Отримуємо всі форми обох імен
    const forms1 = this.getAllFormsOfName(name1);
    const forms2 = this.getAllFormsOfName(name2);

    // Перевіряємо перетин
    const forms1Lower = forms1.map(f => f.toLowerCase());
    const forms2Lower = forms2.map(f => f.toLowerCase());

    return forms1Lower.some(form => forms2Lower.includes(form));
  }

  /**
   * Знайти найкращий збіг імені в тексті (строга версія)
   */
  public static findNameMatchStrict(text: string, targetName: string): boolean {
    const targetForms = this.getAllFormsOfName(targetName);
    
    return targetForms.some(form => {
      const words = form.split(' ');
      return words.every(word => {
        if (word.length < 3) return true; // Ігноруємо короткі слова
        return text.toLowerCase().includes(word.toLowerCase());
      });
    });
  }

  /**
   * Знайти збіг імені в тексті (більш строга версія) - потрібно знайти всі частини імені
   * ВАЖЛИВО: Шукаємо цілі слова, а не частини інших слів
   */
  public static findNameMatch(text: string, targetName: string): boolean {
    const targetForms = this.getAllFormsOfName(targetName);
    const textLower = text.toLowerCase();
    
    // Розбиваємо текст на слова (розділювачі: пробіли, коми, крапки, дужки тощо)
    const textWords = textLower.split(/[\s,.:;()«»"""''\/\\—–-]+/).filter(w => w.length > 0);
    
    // Спробувати знайти точний збіг для кожної форми
    for (const form of targetForms) {
      const formWords = form.split(' ').filter(w => w.length >= 2);
      
      // Перевіряємо, чи всі слова форми є серед окремих слів тексту
      const allWordsFound = formWords.every(formWord => {
        const formWordLower = formWord.toLowerCase();
        return textWords.some(textWord => textWord === formWordLower);
      });
      
      if (formWords.length > 0 && allWordsFound) {
        return true;
      }
    }
    
    return false;
  }

  /**
   * Діагностична версія для відладки - повертає детальну інформацію про пошук
   * ВАЖЛИВО: Шукаємо цілі слова, а не частини інших слів
   */
  public static findNameMatchDebug(text: string, targetName: string): {
    found: boolean;
    matchDetails: Array<{
      form: string;
      matchedWords: string[];
      allWords: string[];
      matchRatio: number;
    }>;
  } {
    const targetForms = this.getAllFormsOfName(targetName);
    const textLower = text.toLowerCase();
    const textWords = textLower.split(/[\s,.:;()«»"""''\/\\—–-]+/).filter(w => w.length > 0);
    
    const matchDetails: Array<{
      form: string;
      matchedWords: string[];
      allWords: string[];
      matchRatio: number;
    }> = [];
    
    let found = false;
    
    for (const form of targetForms) {
      const formWords = form.split(' ').filter(w => w.length >= 2);
      const matchedWords = formWords.filter(formWord => 
        textWords.some(textWord => textWord === formWord.toLowerCase())
      );
      const matchRatio = formWords.length > 0 ? matchedWords.length / formWords.length : 0;
      
      matchDetails.push({
        form,
        matchedWords,
        allWords: formWords,
        matchRatio
      });
      
      // Умови для позитивного результату - потрібно знайти ВСІ слова як окремі слова
      if (formWords.length > 0 && matchedWords.length === formWords.length) {
        found = true;
      }
    }
    
    return { found, matchDetails };
  }

  /**
   * ТЕСТОВИЙ МЕТОД: Перевірити як працює відмінювання конкретного ПІБ
   */
  public static testNameDeclension(fullName: string): {
    originalName: string;
    allForms: string[];
    lastNameForms: string[];
    firstNameForms: string[];
    middleNameForms: string[];
  } {
    const parts = fullName.trim().split(/\s+/);
    
    let lastName: string = '', firstName: string = '', middleName: string = '';
    if (parts.length >= 3) {
      [lastName, firstName, middleName] = parts;
    } else if (parts.length === 2) {
      [lastName, firstName] = parts;
    }

    return {
      originalName: fullName,
      allForms: this.getAllFormsOfName(fullName),
      lastNameForms: this.getLastNameForms(lastName),
      firstNameForms: this.getFirstNameForms(firstName),
      middleNameForms: middleName ? this.getMiddleNameForms(middleName) : []
    };
  }

  /**
   * Обробка Excel колонки D з ПІБ для режиму "Розпорядження"
   * @param excelNames - масив ПІБ з колонки D Excel файлу (у називному відмінку)
   * @param wordText - текст Word документу з ПІБ в різних відмінках
   * @returns масив знайдених збігів з інформацією про позиції
   */
  public static processExcelColumnD(excelNames: string[], wordText: string): Array<{
    excelName: string;
    foundInText: boolean;
    matchedForms: string[];
    positions: number[];
  }> {
    const results: Array<{
      excelName: string;
      foundInText: boolean;
      matchedForms: string[];
      positions: number[];
    }> = [];

    for (const excelName of excelNames) {
      if (!excelName || excelName.trim() === '') continue;

      const cleanName = excelName.trim();
      const matchedForms: string[] = [];
      const positions: number[] = [];

      // Отримуємо всі форми імені з Excel (називний відмінок)
      const allForms = this.getAllFormsOfName(cleanName);

      // Шукаємо кожну форму в тексті Word документу
      for (const form of allForms) {
        const words = form.split(' ').filter(w => w.length >= 3);
        
        for (const word of words) {
          const regex = new RegExp(`\\b${word.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}\\b`, 'gi');
          let match;
          
          while ((match = regex.exec(wordText)) !== null) {
            if (!positions.includes(match.index)) {
              positions.push(match.index);
              matchedForms.push(form);
            }
          }
        }
      }

      results.push({
        excelName: cleanName,
        foundInText: positions.length > 0,
        matchedForms: [...new Set(matchedForms)],
        positions: positions.sort((a, b) => a - b)
      });
    }

    return results;
  }

  /**
   * ВАРІАНТ 1: Знайти тільки абзаци з ключовим словом "розпорядженні" (без урахування ПІБ)
   */
  public static findOrderParagraphsOnly(wordText: string): Array<{
    paragraph: string;
    startPosition: number;
  }> {
    const paragraphs = wordText.split(/\n\s*\n/).filter(p => p.trim());
    const results: Array<{
      paragraph: string;
      startPosition: number;
    }> = [];

    let currentPosition = 0;

    for (const paragraph of paragraphs) {
      const containsOrderKeyword = /розпоряд(женн[іїя]|ження)/i.test(paragraph);

      if (containsOrderKeyword) {
        results.push({
          paragraph: paragraph.trim(),
          startPosition: currentPosition
        });
      }

      currentPosition += paragraph.length + 2; // +2 для \n\n
    }

    return results;
  }

  /**
   * ВАРІАНТ 2: Знайти абзаци з ПІБ з Excel (без урахування "розпорядженні")
   */
  public static findParagraphsWithExcelNames(
    wordText: string, 
    excelNames: string[]
  ): Array<{
    paragraph: string;
    matchedNames: string[];
    startPosition: number;
  }> {
    const paragraphs = wordText.split(/\n\s*\n/).filter(p => p.trim());
    const results: Array<{
      paragraph: string;
      matchedNames: string[];
      startPosition: number;
    }> = [];

    let currentPosition = 0;

    for (const paragraph of paragraphs) {
      const matchedNames: string[] = [];

      // Шукаємо ПІБ з Excel в цьому абзаці
      for (const excelName of excelNames) {
        if (!excelName || excelName.trim() === '') continue;
        
        if (this.findNameMatch(paragraph, excelName.trim())) {
          matchedNames.push(excelName.trim());
        }
      }

      if (matchedNames.length > 0) {
        results.push({
          paragraph: paragraph.trim(),
          matchedNames: [...new Set(matchedNames)],
          startPosition: currentPosition
        });
      }

      currentPosition += paragraph.length + 2; // +2 для \n\n
    }

    return results;
  }

  /**
   * ВАРІАНТ 3: Знайти абзаци з ключовим словом "розпорядженні" та ПІБ з Excel колонки D
   * (Потрібні ОБИДВА критерії)
   */
  public static findOrderParagraphs(
    wordText: string, 
    excelNames: string[]
  ): Array<{
    paragraph: string;
    containsOrderKeyword: boolean;
    matchedNames: string[];
    startPosition: number;
  }> {
    const paragraphs = wordText.split(/\n\s*\n/).filter(p => p.trim());
    const results: Array<{
      paragraph: string;
      containsOrderKeyword: boolean;
      matchedNames: string[];
      startPosition: number;
    }> = [];

    let currentPosition = 0;

    for (const paragraph of paragraphs) {
      const containsOrderKeyword = /розпоряд(женн[іїя]|ження)/i.test(paragraph);
      const matchedNames: string[] = [];

      if (containsOrderKeyword) {
        // Шукаємо ПІБ з Excel в цьому абзаці
        for (const excelName of excelNames) {
          if (!excelName || excelName.trim() === '') continue;
          
          if (this.findNameMatch(paragraph, excelName.trim())) {
            matchedNames.push(excelName.trim());
          }
        }
      }

      // ЗМІНА: повертаємо тільки абзаци, які містять І "розпорядженні" І ПІБ з Excel
      if (containsOrderKeyword && matchedNames.length > 0) {
        results.push({
          paragraph: paragraph.trim(),
          containsOrderKeyword: true,
          matchedNames: [...new Set(matchedNames)],
          startPosition: currentPosition
        });
      }

      currentPosition += paragraph.length + 2; // +2 для \n\n
    }

    return results;
  }

  /**
   * ДІАГНОСТИЧНА ВЕРСІЯ: Знайти абзаци з "розпорядженні" та ПІБ + детальний лог
   */
  public static findOrderParagraphsDebug(
    wordText: string, 
    excelNames: string[]
  ): {
    results: Array<{
      paragraph: string;
      containsOrderKeyword: boolean;
      matchedNames: string[];
      startPosition: number;
    }>;
    diagnostics: Array<{
      paragraphIndex: number;
      containsOrderKeyword: boolean;
      checkedNames: Array<{
        name: string;
        found: boolean;
        matchDetails: any;
      }>;
      finallyIncluded: boolean;
    }>;
  } {
    const paragraphs = wordText.split(/\n\s*\n/).filter(p => p.trim());
    const results: Array<{
      paragraph: string;
      containsOrderKeyword: boolean;
      matchedNames: string[];
      startPosition: number;
    }> = [];
    
    const diagnostics: Array<{
      paragraphIndex: number;
      containsOrderKeyword: boolean;
      checkedNames: Array<{
        name: string;
        found: boolean;
        matchDetails: any;
      }>;
      finallyIncluded: boolean;
    }> = [];

    let currentPosition = 0;

    for (let i = 0; i < paragraphs.length; i++) {
      const paragraph = paragraphs[i];
      const containsOrderKeyword = /розпоряд(женн[іїя]|ження)/i.test(paragraph);
      const matchedNames: string[] = [];
      const checkedNames: Array<{
        name: string;
        found: boolean;
        matchDetails: any;
      }> = [];

      if (containsOrderKeyword) {
        // Шукаємо ПІБ з Excel в цьому абзаці
        for (const excelName of excelNames) {
          if (!excelName || excelName.trim() === '') continue;
          
          const nameDebug = this.findNameMatchDebug(paragraph, excelName.trim());
          const found = nameDebug.found;
          
          checkedNames.push({
            name: excelName.trim(),
            found,
            matchDetails: nameDebug.matchDetails
          });
          
          if (found) {
            matchedNames.push(excelName.trim());
          }
        }
      }

      // ВИПРАВЛЕНО: додаємо абзац тільки якщо В ОДНОМУ абзаці є І ПІБ з Excel І "розпорядженні"
      const finallyIncluded = containsOrderKeyword && matchedNames.length > 0;
      
      if (finallyIncluded) {
        results.push({
          paragraph: paragraph.trim(),
          containsOrderKeyword: true,
          matchedNames: [...new Set(matchedNames)],
          startPosition: currentPosition
        });
      }

      diagnostics.push({
        paragraphIndex: i,
        containsOrderKeyword,
        checkedNames,
        finallyIncluded
      });

      currentPosition += paragraph.length + 2; // +2 для \n\n
    }

    return { results, diagnostics };
  }

  /**
   * ВАРІАНТ 4: Комплексний аналіз - всі абзаци з детальною інформацією
   */
  public static analyzeAllParagraphs(
    wordText: string, 
    excelNames: string[]
  ): Array<{
    paragraph: string;
    containsOrderKeyword: boolean;
    matchedNames: string[];
    startPosition: number;
    shouldInclude: boolean; // чи включати цей абзац в результат
  }> {
    const paragraphs = wordText.split(/\n\s*\n/).filter(p => p.trim());
    const results: Array<{
      paragraph: string;
      containsOrderKeyword: boolean;
      matchedNames: string[];
      startPosition: number;
      shouldInclude: boolean;
    }> = [];

    let currentPosition = 0;

    for (const paragraph of paragraphs) {
      const containsOrderKeyword = /розпоряд(женн[іїя]|ження)/i.test(paragraph);
      const matchedNames: string[] = [];

      // Завжди шукаємо ПІБ, незалежно від "розпорядженні"/"розпорядження"
      for (const excelName of excelNames) {
        if (!excelName || excelName.trim() === '') continue;
        
        if (this.findNameMatch(paragraph, excelName.trim())) {
          matchedNames.push(excelName.trim());
        }
      }

      // ВИПРАВЛЕНО: Логіка включення - І "розпорядженні" І ПІБ з Excel в одному абзаці
      const shouldInclude = containsOrderKeyword && matchedNames.length > 0;

      results.push({
        paragraph: paragraph.trim(),
        containsOrderKeyword,
        matchedNames: [...new Set(matchedNames)],
        startPosition: currentPosition,
        shouldInclude
      });

      currentPosition += paragraph.length + 2; // +2 для \n\n
    }

    return results;
  }
}

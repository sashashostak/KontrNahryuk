# 🧹 Інструкція з очищення Git репозиторію від великих файлів

## Проблема
Папка `release/` (911 MB, 4,811 файлів) була випадково додана до Git історії.  
Хоча вона зараз в `.gitignore`, старі commits все ще містять ці файли.

## Статус
- ✅ `.gitignore` налаштовано правильно
- ✅ Нові commits не будуть включати `release/`
- ⚠️ Старі commits все ще містять великі файли

## Опція 1: Залишити як є (рекомендовано для невеликих команд)
Якщо ви працюєте один або в невеликій команді, можна просто продовжувати роботу.  
Нові commits будуть легкими.

## Опція 2: Очистити історію (для великих команд / публічних репозиторіїв)

### Використання BFG Repo-Cleaner (найпростіше)

```bash
# 1. Завантажити BFG
# https://rtyley.github.io/bfg-repo-cleaner/

# 2. Створити backup
git clone --mirror https://github.com/sashashostak/KontrNahryuk.git backup-repo.git

# 3. Видалити папку release/ з історії
java -jar bfg.jar --delete-folders release --no-blob-protection KontrNahryuk.git

# 4. Очистити Git
cd KontrNahryuk.git
git reflog expire --expire=now --all
git gc --prune=now --aggressive

# 5. Force push (УВАГА: перепише історію!)
git push --force
```

### Використання git filter-repo (альтернатива)

```bash
# 1. Встановити git-filter-repo
pip install git-filter-repo

# 2. Видалити папку release/
git filter-repo --path release/ --invert-paths

# 3. Force push
git push --force
```

## ⚠️ ВАЖЛИВО при force push:
1. Попередити всіх членів команди
2. Всі мають зробити `git pull --rebase` після вашого push
3. Краще робити в неробочий час

## Альтернатива: Створити новий репозиторій
Якщо очищення історії складне:
1. Створити новий GitHub репозиторій
2. Скопіювати тільки вихідний код (без release/)
3. Зробити initial commit
4. Старий репозиторій архівувати

## Перевірка розміру репозиторію

```bash
# Розмір .git папки
du -sh .git

# Найбільші файли в історії
git rev-list --objects --all | \
  git cat-file --batch-check='%(objecttype) %(objectname) %(objectsize) %(rest)' | \
  awk '/^blob/ {print substr($0,6)}' | \
  sort --numeric-sort --key=2 | \
  cut --complement --characters=13-40 | \
  numfmt --field=2 --to=iec-i --suffix=B --padding=7 --round=nearest | \
  tail -20
```

## Рекомендація
Для цього проекту рекомендую **Опцію 1** - залишити як є.  
Нові commits будуть легкими, а історію можна очистити пізніше при потребі.

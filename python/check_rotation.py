import csv

data = list(csv.reader(open('test_real_zbd.csv', encoding='utf-8-sig')))

print('Події з ротацією (вибули):')
for i, row in enumerate(data[1:]):
    for j in range(3, min(len(row), 34)):
        if row[j].strip() and 'вибул' in row[j].lower():
            print(f'\nРядок {i+2}, Колонка {j+1} (День {j-2}):')
            print(repr(row[j]))
            print('='*50)

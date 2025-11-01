import csv

data = list(csv.reader(open('test_real_zbd.csv', encoding='utf-8-sig')))

print('Події з часом 07:00:')
for i, row in enumerate(data[1:]):
    for j in range(3, min(len(row), 34)):
        if row[j].strip() and '07:00' in row[j]:
            print(f'\nРядок {i+2}, Колонка {j+1} (День {j-2}):')
            print(row[j])
            print('='*60)

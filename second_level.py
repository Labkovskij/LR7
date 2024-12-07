import pandas as pd
import numpy as np

# Завантаження даних з Excel файлу
file_path = 'Data_Set_6.xlsx'
df = pd.read_excel(file_path, engine='openpyxl')

# Попередній аналіз даних
print("Опис даних:")
print(df.describe())  # статистичний опис даних

print("\nПеревірка на пропущені значення:")
print(df.isnull().sum())  # кількість пропущених значень в кожному стовпці

print("\nІнформація про дані:")
print(df.info())  # інформація про типи даних

# Очистка даних: Видалення рядків з пропущеними значеннями
df_cleaned = df.dropna()

# Перетворення числових стовпців на числові значення (виправлення типу для нечислових значень)
numeric_columns = df_cleaned.select_dtypes(include=[np.number]).columns

# Оцінка кореляцій між змінними
print("\nКореляція між числовими змінними:")
correlation_matrix = df_cleaned[numeric_columns].corr()
print(correlation_matrix)

# Додатковий аналіз згідно з технічними умовами
# Приклад аналізу: розподіл значень по категоріям або групам
# Наприклад, якщо є категоріальний стовпець 'SALES_BY_REGION', можна виконати агрегацію:

if 'SALES_BY_REGION' in df.columns:
    category_analysis = df_cleaned.groupby('SALES_BY_REGION').agg({
        'APRIL': ['mean', 'std', 'min', 'max'],
        'MAY': ['mean', 'std', 'min', 'max'],
        'SEPTEMBER': ['mean', 'std', 'min', 'max'],
        'NOVEMBER': ['mean', 'std', 'min', 'max']
    })
    print("\nАналіз за категоріями регіонів:")
    print(category_analysis)

# Виведення основних статистичних результатів
print("\nОсновні статистичні результати:")
print(df_cleaned.describe())

# Збереження очищених даних у новий файл
df_cleaned.to_excel('cleaned_data.xlsx', index=False)

# Візуалізація результатів (наприклад, гістограма для певного стовпця)
import matplotlib.pyplot as plt

if 'APRIL' in df_cleaned.columns:
    df_cleaned['APRIL'].hist(bins=20)
    plt.title('Розподіл значень по квітню')
    plt.xlabel('Значення')
    plt.ylabel('Частота')
    plt.show()

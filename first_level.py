import pandas as pd
import numpy as np
from sklearn.linear_model import LinearRegression
import matplotlib.pyplot as plt

# Завантаження Excel файлу
df = pd.read_excel('Pr_1.xls', engine='xlrd')

# Попередній аналіз даних
print(df.describe())
print(df.isnull().sum())

# Обчислення продажів та прибутку
df['sales'] = df['Ціна реализації'] * df['КільКість реалізацій']
df['profit'] = df['sales'] - df['Собівартість одиниці']
print(df[['sales', 'profit']].head())

# Видалення рядків з пропущеними значеннями у 'Місяц' та 'sales'
df = df.dropna(subset=['Місяц', 'sales'])

# Перетворення місяців на числа
month_mapping = {
    'Січень': 1, 'Лютий': 2, 'Березень': 3, 'Квітень': 4, 'Травень': 5, 'Червень': 6,
    'Липень': 7, 'Серпень': 8, 'Вересень': 9, 'Жовтень': 10, 'Листопад': 11, 'Грудень': 12
}

# Перетворення місяців на числа
df['Місяц'] = df['Місяц'].map(month_mapping)

# Видалення рядків з пропущеними значеннями в колонці 'Місяц' після перетворення
df = df.dropna(subset=['Місяц'])

# Математичне моделювання
X = df[['Місяц']]  # незалежна змінна
y = df['sales']    # залежна змінна
model = LinearRegression()
model.fit(X, y)

# Прогнозування
months_future = np.array([df['Місяц'].max() + i for i in range(1, 7)]).reshape(-1, 1)
sales_forecast = model.predict(months_future)

# Графік
plt.plot(df['Місяц'], df['sales'], label='Історичні дані')
plt.plot(months_future, sales_forecast, label='Прогноз', linestyle='--')
plt.xlabel('Місяць')
plt.ylabel('Продажі')
plt.legend()
plt.show()

# Результати
print("Прогноз продажів на наступні 6 місяців:", sales_forecast)

import pandas, math
import matplotlib.pyplot as plt
from datetime import datetime

# Загрузка данных из файла Excel
df = pandas.read_excel('data.xlsx',index_col=False)

# Инициализация данных менеджеров
managers = {}
original_contracts_count = 0
# Рассматриваемая дата
target_date = datetime(2021, 7, 1)
df['receiving_date'] = pandas.to_datetime(df['receiving_date'], format='%Y-%m-%d',errors='coerce')
total_july = 0
# Вычисление бонусов для каждой сделки
deal_date = datetime(2021, 5,1)
for index, row in df.iterrows():
    try:
        df.at[index, 'receiving_date'] = pandas.to_datetime(row['receiving_date'])
    except:
        pass
    if row['sale'] != '-':
        manager = row['sale']
    else: continue
    if '2021' in row['status']:
        if 'июнь' in row['status'].lower():
            deal_date = datetime(2021, 6, 1)
        if 'июль' in row['status'].lower():
            deal_date = datetime(2021, 7, 1)
        if 'август' in row['status'].lower():
            deal_date = datetime(2021, 8, 1)
        if 'сентябрь' in row['status'].lower():
            deal_date = datetime(2021, 9, 1)
        if 'октябрь' in row['status'].lower():
            deal_date = datetime(2021, 10, 1)
    df.at[index,'deal_date'] = deal_date
    if deal_date < target_date and deal_date < row['receiving_date']:
        if row['new/current'] == 'новая' and row['status'] == 'ОПЛАЧЕНО' and row['document'] == 'оригинал':
            bonus = row['sum'] * 0.07
            managers[manager] = managers.get(manager, 0) + bonus
        if row['new/current'] == 'текущая' and row['status'] != 'ПРОСРОЧЕНО' and row['document'] == 'оригинал':
            if row['sum'] > 10000:
                bonus = row['sum'] * 0.05
            else:
                bonus = row['sum'] * 0.03
            managers[manager] = managers.get(manager, 0) + bonus
    if deal_date == datetime(2021, 7, 1):
        if not math.isnan(row['sum']) and row['status'] != 'ПРОСРОЧЕНО':
            total_july += row['sum']
    if deal_date.month == 5 and row['receiving_date'].month == 6 and row['document'] == 'оригинал':
        original_contracts_count += 1

# Вывод остатков бонусов на 01.07.2021 длякаждого менеджера
for manager, bonus in managers.items():
    print(f'{manager}: {round(bonus,2)}')

# Общая выручка за июль 2021 года по непросроченным сделкам
print(f'Общая выручка за июль 2021: {round(total_july,2)}')
#print(df)
# Наиболее успешный менеджер по привлечениюденежных средств в сентябре 2021
max_revenue_manager_sept = df[(df['deal_date'] >= datetime.strptime('2021-09-01', '%Y-%m-%d')) &
                              (df['deal_date'] <= datetime.strptime('2021-09-30', '%Y-%m-%d'))]['sale'].value_counts().idxmax()
print(f'Наиболее успешный менеджер в сентябре2021: {max_revenue_manager_sept}')

# Преобладающий тип сделок (новая/текущая) воктябре 2021
deal_type_oct = df[(df['deal_date'] >= datetime.strptime('2021-10-01', '%Y-%m-%d')) &
                   (df['deal_date'] <= datetime.strptime('2021-10-31', '%Y-%m-%d'))]['new/current'].value_counts().idxmax()
print(f'Преобладающий тип сделок в октябре2021: {deal_type_oct}')

print(f'Количество оригиналов договоров помайским сделкам в июне 2021: {original_contracts_count}')


# Вычисление изменения выручки компании зарассматриваемые периоды и построение графика
revenue_by_month = df.groupby(df['deal_date'].dt.to_period('M'))['sum'].sum()
revenue_by_month.plot(kind='bar', title='Выручка компании по месяцам')
plt.xlabel('Месяц')
plt.ylabel('Выручка')
plt.show()
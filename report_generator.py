import pandas as pd

df = pd.read_csv("sales.csv")
print(df.info)



if 'revenue' not in df.columns:
    df["revenue"] = df["price"] * df["quantity"]

df['date_dt'] = pd.to_datetime(df['date'])
df['month'] = df['date_dt'].dt.strftime('%B')  
df['weekday'] = df['date_dt'].dt.day_name()    

df['date'] = df['date_dt'].dt.date
df = df.drop(columns=['date_dt'])

'''Общая выручка'''
total_revenue = df["revenue"].sum()
'''Средний чек'''
avg_order = df["revenue"].mean()
'''Топ продукты'''
top_products = df.groupby("product")["revenue"].sum().sort_values(ascending=False).head(5)



'''метрики'''

total_quantity = df["quantity"].sum()
avg_price = df["price"].mean()
best_selling_product = df.groupby("product")["quantity"].sum().sort_values(ascending=False).head(1).index[0]
monthly_revenue = df.groupby("month")["revenue"].sum()


print(f"Всего строк: {len(df)}")
print(f"Всего стобцов: {list(df.columns)}")
print(f"Общая выручка: {df['revenue'].sum():,.0f} ₽")
print(f"Всего продано:  {total_quantity:,} шт")
print(f"Средний чек: {avg_order:,.0f} ₽")
print(f"Средняя цена товара: {avg_price:,.0f} ₽")
print(f"Самый продаваемый: {best_selling_product}")
print(f"Уникальных товаров: {df['product'].nunique()}")
print(f"Регионы: {df['region'].unique()}")
print(f"Всего возвратов: {df['returned'].sum()}")

with pd.ExcelWriter("weekly_report.xlsx") as writer:
    df.to_excel(writer, sheet_name="Данные", index=False)
    
    summary = pd.DataFrame({
        "Метрика": [
            "Общая выручка",
            "Средний чек",
            "Всего продано единиц",
            "Лучший продукт"
        ],
        "Значение": [
            f"{total_revenue:,.0f} ₽",
            f"{avg_order:,.0f} ₽",
            f"{df['quantity'].sum():,}",
            top_products.index[0] if len(top_products) > 0 else "-"
        ]
    })
    summary.to_excel(writer, sheet_name="Сводка", index=False)
    
    # Лист 3: топ продуктов (исправлено!)
    top_products_df = pd.DataFrame({
        "Продукт": top_products.index,
        "Выручка": [f"{val:,.0f} ₽" for val in top_products.values] 
    })
    top_products_df.to_excel(writer, sheet_name="Топ продуктов", index=False)

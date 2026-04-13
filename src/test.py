import pandas as pd
import matplotlib.pyplot as plt

data = pd.read_excel('superstore_sales.xlsx')
#print(data.to_string())

print("Number of missing values by column is: ")
print(data.isnull().sum())

#Cleaning the duplicates and empty fields
data.drop_duplicates()
data['profit'] = data['profit'].fillna(0)

#Grouping the sales data by month/year
data['month_year'] = data['order_date'].apply(lambda x: x.strftime('%Y-%m'))
data_trend = data.groupby('month_year').sum(numeric_only=True)['sales'].reset_index()
print(data_trend)

#Visualizing the data in matplot
plt.figure(figsize=(15,6))
plt.plot(data_trend['month_year'], data_trend['sales'])
plt.xticks(rotation='vertical')
plt.show()

#Highlighting top 10 products by sales
sales = pd.DataFrame(data.groupby('product_name').sum(numeric_only=True)['sales'])
sales = sales.sort_values('sales', ascending=False)
print(sales[:10])

#Highlighting the most profitable categories
profit = pd.DataFrame(data.groupby(['category', 'sub_category']).sum(numeric_only=True)['profit'])
profit = profit.sort_values('profit', ascending=False)
print(profit)

#Highlighting the dead stock in last 3 months
last_date = pd.to_datetime(data['month_year'].max())
start = pd.to_datetime(last_date - pd.DateOffset(months=3))
data['month_year'] = pd.to_datetime(data['month_year'])
quarter = data[(data['month_year'] >= start) & (data['month_year'] < last_date)].groupby(['month_year', 'sales', 'product_name']).sum(numeric_only=True)
quarter = quarter.sort_values('sales', ascending=True)
print(quarter[:10])

#Pivot table
pivot = data.pivot_table(index='market', columns='category', values='profit', aggfunc='sum')
print(pivot)
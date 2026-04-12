import pandas as pd
import matplotlib.pyplot as plt

data = pd.read_excel('superstore_sales.xlsx')
#print(data.to_string())

print("Number of missing values by column is: ")
print(data.isnull().sum())

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
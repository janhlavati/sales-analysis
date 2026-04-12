import pandas as pd
import matplotlib.pyplot as plt

data = pd.read_excel('superstore_sales.xlsx')
#print(data.to_string())

print("Number of missing values by column is: ")
print(data.isnull().sum())

#Grouping the sales data by month/year
data['month_year'] = data['order_date'].apply(lambda x: x.strftime('%Y-%m'))
data_trend = data.groupby('month_year').sum(numeric_only = True)['sales'].reset_index()
print(data_trend)

#Visualizing the data in matplot
plt.figure(figsize=(15,6))
plt.plot(data_trend['month_year'], data_trend['sales'])
plt.xticks(rotation='vertical')
plt.show()
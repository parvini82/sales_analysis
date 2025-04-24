import pandas as pd
import openpyxl as xl
import matplotlib.pyplot as plt

# Load Excel dataset
df = pd.read_excel('/Users/work/PycharmProjects/sales_analysis/data/Retail.xlsx')

# Replace missing CustomerID with -1 and convert to integer32
df['CustomerID'] = df['CustomerID'].fillna(-1).astype('int32')
# Fill missing product descriptions with "NONE"
df['Description'] = df['Description'].fillna('NONEٔ')

# Separate returns and actual sales
df_returns = df[df['Quantity'] < 0].copy()
df_sales = df[df['Quantity'] > 0].copy()

# Convert InvoiceDate to datetime format
df_sales.loc[:, 'InvoiceDate'] = pd.to_datetime(df_sales['InvoiceDate'])

# Calculate TotalPrice as Quantity * UnitPrice
df_sales.loc[:, 'TotalPrice'] = df_sales['Quantity'] * df_sales['UnitPrice']

# Set InvoiceDate as index for time-based analysis
df_sales.set_index('InvoiceDate', inplace=True)

# Overall sales analysis
total_sales = df_sales['TotalPrice'].sum()
total_transactions = df_sales['InvoiceNo'].nunique()
average_order_value = total_sales / total_transactions

# Monthly sales trend
monthly_sales = df_sales['TotalPrice'].resample('ME').sum()

# Sales by country
total_sales_by_country = df_sales.groupby('Country')['TotalPrice'].sum()
max_sales_by_country = total_sales_by_country.sort_values(ascending=False).head(3)

# Top-selling products by revenue
max_sales_by_product = df_sales.groupby('Description')['TotalPrice'].sum().sort_values(ascending=False).head(3)

# Top-selling products by quantity
top_quantities = df_sales.groupby('Description')['Quantity'].sum().sort_values(ascending=False).head(3)

# Total purchase amount per customer (excluding -1)
customer_total = df_sales[df_sales['CustomerID'] != -1].groupby('CustomerID')['TotalPrice'].sum().sort_values(ascending=False)

monthly_country_sales = df_sales.reset_index().pivot_table(
    values='TotalPrice',
    index=df_sales.index.to_period('M'),  # Monthly index
    columns='Country',
    aggfunc='sum'
)
print(monthly_country_sales.head())


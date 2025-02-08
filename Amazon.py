import numpy as np
import pandas as pd

sales_data = pd.read_excel('sales_data.xlsx')

# Exploring the data

# Get a summary of sales data
sales_data.info()

sales_data.describe()

# Looking at columns
print(sales_data.columns)

# Having a look at the first few rows of data
print(sales_data.head())

# Check the data types of the columns
print(sales_data.dtypes)


# Cleaning the data

# Check for missing values in our sales data
sales_data.isnull()

# Drop any rows that has any missing/nan values
sales_data_dropped = sales_data.dropna()

# Drop rows with missing amounts based on amount column
sales_data_cleaned = sales_data.dropna(subset = ['Amount'])

# Check for missing values in our sales data cleaned
print(sales_data_cleaned.isnull().sum())


# Aggregating data

# Toal sales by category
category_totals = sales_data.groupby('Category')['Amount'].sum()
category_totals = sales_data.groupby('Category', as_index=False)['Amount'].sum()
category_totals = category_totals.sort_values('Amount', ascending=False)

# Calculate the average Amount by category and fulfilment
fulfilment_averages = sales_data.groupby(['Category', 'Fulfilment'], as_index=False)['Amount'].mean()
fulfilment_averages = fulfilment_averages.sort_values('Amount', ascending=False)

# Calculate the average Amount by category and status
status_averages = sales_data.groupby(['Category', 'Status'], as_index=False)['Amount'].mean()
status_averages = status_averages.sort_values('Amount', ascending=False)

# Calculate total sales by shipment and fulfilment
total_sales_shipandfulfil = sales_data.groupby(['Courier Status', 'Fulfilment'], as_index=False)['Amount'].mean()
total_sales_shipandfulfil = total_sales_shipandfulfil.sort_values('Amount', ascending=False)
total_sales_shipandfulfil.rename(columns={'Courier Status' : 'Shipment'}, inplace=True)


# Exporting the Data

status_averages.to_excel('average_sales_by_category_and_status.xlsx', index=False)
total_sales_shipandfulfil.to_excel('toal_sales_by_ship_and_fulfil.xlsx', index=False)
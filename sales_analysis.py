# IMPORTING REQUIRED LIBRARIES
# Data Manipulation
import pandas as pd

# Data Visualisation
import matplotlib.pyplot as plt
import seaborn as sns

# IMPORTING THE DATASET
# Importing dataset
df = pd.read_excel('superstore_sales.xlsx')

# First five rows of the dataset
print(df.head())

# Last five rows of the dataset
print(df.tail())

# Qty Number of the dataset and column
print(df.shape)

# Columns present in the dataset
print(df.columns)

# A concise summary of the dataset
print(df.info())

# Checking missing values
print(df.isna().sum())

# Generating descriptive statistics summary
print(df.describe().round())


# EXPLORATORY DATA ANALYSIS

# 1. WHAT IS THE OVERALL SALES TREND?
# Getting month year from order_date
df['month_year'] = df['order_date'].apply(lambda x: x.strftime('%m-%Y'))
print(df['month_year'])

# grouping month_year by sales
df_temp = df.groupby('month_year').sum()['sales'].reset_index()
print(df_temp)

# Setting the figure size
plt.figure(figsize=(16, 5))
plt.plot(df_temp['month_year'], df_temp['sales'], color='#b80045')
plt.xticks(rotation='vertical', size=8)
# plt.show()


# 2. WHICH ARE THE TOP 10 PRODUCTS BY SALES?
# Grouping products by sales
prod_sales = pd.DataFrame(df.groupby('product_name').sum()['sales'])
# print(prod_sales)

# Sorting the dataframe in descending order
prod_sales.sort_values(by=['sales'], inplace=True, ascending=False)
print(prod_sales)

# Top 10 products by sales
top_prod_sales = prod_sales[:10]
print(top_prod_sales)


# 3. WHICH ARE THE MOST SELLING PRODUCTS?
# Grouping products by Quantity
best_selling_prods = pd.DataFrame(df.groupby('product_name').sum()['quantity'])
# print(best_selling_prods)

# Sorting the dataframe in descending order
best_selling_prods.sort_values(by=['quantity'], inplace=True, ascending=False)
print(best_selling_prods)

# Most selling products
most_best_selling_prods = best_selling_prods[:10]
print(most_best_selling_prods)


# 4. WHAT IS THE MOST PREFERRED SHIP MODE?
# Setting the figure size
plt.figure(figsize=(10, 8))
# countplot: Show the counts of observations in each categorical bin using bars
sns.countplot(x='ship_mode', data=df)
# Display the figure
# plt.show()


# 5. WHICH ARE THE MOST PROFITABLE CATEGORY AND SUB-CATEGORY?
# Grouping products by Category and Sub-Category
cat_subcat = pd.DataFrame(df.groupby(['category', 'sub_category']).sum()['profit'])
# print(cat_subcat)

# Sorting the values
cat_subcat.sort_values(by=['category', 'profit'], inplace=True, ascending=False)
print(cat_subcat)

import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from datetime import datetime
import warnings

warnings.filterwarnings('ignore')

data = pd.read_csv(r"C:\Users\tmrma\OneDrive\Desktop\Ecommerce_Delivery_Analytics_New.csv")
data.columns = data.columns.str.strip()
print("Column Names in the Dataset (after stripping whitespace):")
print(data.columns.tolist())

date_column = 'Order Date'
if date_column not in data.columns:
    possible_names = ['OrderDate', 'order_date', 'Order date', 'order Date']
    for name in possible_names:
        if name in data.columns:
            date_column = name
            print(f"Found date column: {date_column}")
            break
    else:
        raise KeyError("No 'Order Date' column or variant found in the dataset. Please check your CSV file.")

data[date_column] = pd.to_datetime('2025-03-11 ' + data[date_column], format='%Y-%m-%d %H:%M:%S.%f', errors='coerce')

unique_orders = data['Order ID'].nunique()
customer_order_counts = data['Customer ID'].value_counts()
frequent_customers = customer_order_counts[customer_order_counts > 10]

print(f"Total Unique Orders: {unique_orders}")
print("Customers with >10 Orders:")
print(frequent_customers)

platform_group = data.groupby('Platform')
category_sales = data.groupby(['Platform', 'Product Category'])['Order ID'].nunique().unstack().fillna(0)

category_sales.plot(kind='bar', stacked=True, figsize=(10, 6))
plt.title('Unique Sales by Product Category Across Platforms')
plt.xlabel('Platform')
plt.ylabel('Number of Unique Orders')
plt.legend(title='Product Category')
plt.tight_layout()
plt.savefig('category_sales_by_platform.png')
plt.show()

data['Year'] = data['Order Date'].dt.year
data['Month'] = data['Order Date'].dt.month
data['Week'] = data['Order Date'].dt.isocalendar().week

yearly_orders = data.groupby('Year')['Order ID'].nunique()
print("Yearly Orders:")
print(yearly_orders)

monthly_orders = data.groupby('Month')['Order ID'].nunique()
print("Monthly Orders:")
print(monthly_orders)

weekly_orders = data.groupby('Week')['Order ID'].nunique()
print("Weekly Orders:")
print(weekly_orders)

monthly_orders.plot(kind='bar', figsize=(8, 5))
plt.title('Monthly Order Analysis')
plt.xlabel('Month')
plt.ylabel('Number of Orders')
plt.savefig('monthly_orders.png')
plt.show()

platform_performance = data.groupby('Platform').agg({
    'Service Rating': 'mean',
    'Order Value (INR)': 'sum',
    'Order ID': 'nunique'
}).rename(columns={'Order ID': 'Unique Orders'})
print("Platform Performance:")
print(platform_performance)

best_platform = platform_performance['Service Rating'].idxmax()
worst_platform = platform_performance['Service Rating'].idxmin()
print(f"Best Performing Platform: {best_platform}")
print(f"Least Performing Platform: {worst_platform}")

top_rated = data.groupby('Product Category')['Service Rating'].mean().nlargest(10)
print("Top 10 Highest Rated Product Categories:")
print(top_rated)

lowest_rated = data.groupby('Product Category')['Service Rating'].mean().nsmallest(5)
print("Top 5 Lowest Rated Product Categories:")
print(lowest_rated)

profit_analysis = data.groupby('Platform')['Order Value (INR)'].sum()
print("Profit Analysis (Total Order Value in INR):")
print(profit_analysis)

profit_analysis.plot(kind='bar', figsize=(8, 5))
plt.title('Total Order Value by Platform')
plt.xlabel('Platform')
plt.ylabel('Total Order Value (INR)')
plt.savefig('profit_analysis.png')
plt.show()

data['Status'] = 'Delivered'
data.loc[data['Delivery Delay'] == 'Yes', 'Status'] = 'Delayed'
data.loc[data['Refund Requested'] == 'Yes', 'Status'] = 'Cancelled'

status_counts = data['Status'].value_counts(normalize=True) * 100
print("Order Status (%):")
print(status_counts)

status_counts.plot(kind='pie', autopct='%1.1f%%', figsize=(6, 6))
plt.title('Order Status Distribution')
plt.savefig('order_status.png')
plt.show()

print("\nImprovement Suggestions:")
for platform in data['Platform'].unique():
    platform_data = data[data['Platform'] == platform]
    avg_rating = platform_data['Service Rating'].mean()
    delay_rate = (platform_data['Delivery Delay'] == 'Yes').mean() * 100
    refund_rate = (platform_data['Refund Requested'] == 'Yes').mean() * 100
    common_complaints = platform_data[platform_data['Service Rating'] < 3]['Customer Feedback'].value_counts().head(3)

    print(f"\n{platform}:")
    print(f" - Average Rating: {avg_rating:.2f}")
    print(f" - Delay Rate: {delay_rate:.2f}%")
    print(f" - Refund Rate: {refund_rate:.2f}%")
    print(" - Common Complaints:")
    print(common_complaints)
    print(
        " - Suggestion: Improve delivery speed and address frequent issues like missing items or rude delivery personnel.")

with pd.ExcelWriter('ecommerce_analysis.xlsx') as writer:
    data.to_excel(writer, sheet_name='Raw Data', index=False)
    frequent_customers.to_excel(writer, sheet_name='Frequent Customers')
    platform_performance.to_excel(writer, sheet_name='Platform Performance')
    top_rated.to_excel(writer, sheet_name='Top Rated Products')
    lowest_rated.to_excel(writer, sheet_name='Lowest Rated Products')
    profit_analysis.to_excel(writer, sheet_name='Profit Analysis')
    status_counts.to_excel(writer, sheet_name='Order Status')

print("Analysis saved to 'ecommerce_analysis.xlsx'")
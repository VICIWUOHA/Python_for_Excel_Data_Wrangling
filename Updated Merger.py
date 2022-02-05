import pandas as pd

# This Code is used to combine multiple excel files of the same structure by appending them using pandas,
# Its capability also allows declaration on multiple sheets and merging into multiple sheets within the same workbook.
# The 'Openpyxl' package is necessary for implementation alongside pandas

url = [
    r"C:\Users\DATA ANALYST\Python_for_Excel_Data_Wrangling\Data Files\Transactions_1st_January_2013.xlsx",
    r"C:\Users\DATA ANALYST\Python_for_Excel_Data_Wrangling\Data Files\Transactions_2nd_January_2013.xlsx"
]


# initiate empty dataframe which will be used to loop through files
city_orders = pd.DataFrame()
product_orders = pd.DataFrame()

# preprocessing of new data
for file in url:
    city_order_data = pd.read_excel(file, sheet_name="City Orders")
    city_orders = city_orders.append(city_order_data, ignore_index=True)

    product_order_data = pd.read_excel(file, sheet_name='Product Orders')
    product_orders = product_orders.append(product_order_data, ignore_index=True)


# Assign Indexes to both dataframes
city_orders.set_index('Order_Date', inplace=True)

product_orders.set_index('Order_Date', inplace=True)

# write both dataframes to one workbook named 'Merged Sales Data' on two different sheets

with pd.ExcelWriter('Merged Sales Data.xlsx') as writer:
    city_orders.to_excel(writer, sheet_name='City Orders')
    product_orders.to_excel(writer, sheet_name='Product Orders')

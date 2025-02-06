# Provides functions for interacting with the operating system (e.g., file paths, directory creation)
import os
# Provides access to command-line arguments and system-specific functions
import sys 
# A powerful library for data manipulation and analysis
import pandas as pd \
# Provides functions to work with dates and times
from datetime import datetime

def main():
    # Step 1: Get the path to the sales data CSV file
    sales_csv = get_sales_csv()
    
    # Step 2: Create the orders directory
    orders_dir = create_orders_dir(sales_csv)
    
    # Step 3: Process the sales data and generate order files
    process_sales_data(sales_csv, orders_dir)

def get_sales_csv():
    # Check if the correct number of command line arguments is provided
    if len(sys.argv) != 2:
        print("Error: Please provide the path to the sales data CSV file..Exiting...")
        sys.exit(1)
    
    # Get the path from the command line argument
    sales_csv = sys.argv[1]
    
    # Check if the file exists
    if not os.path.isfile(sales_csv):
        print(f"Error: The file '{sales_csv}' does not exist..Exiting...")
        
        sys.exit(1)
    
    return sales_csv

def create_orders_dir(sales_csv):
    # Get the directory of the sales CSV file
    sales_dir = os.path.dirname(sales_csv)
    
    # Generate a date string for the orders directory name
    date_str = datetime.now().strftime("%Y-%m-%d")
    
    # Create the orders directory path
    orders_dir = os.path.join(sales_dir, f"Orders_{date_str}")
    
    # Create the directory if it doesn't exist
    if not os.path.exists(orders_dir):
        os.makedirs(orders_dir)
    
    return orders_dir

def process_sales_data(sales_csv, orders_dir):
    # Read the sales data CSV file into a DataFrame
    df = pd.read_csv(sales_csv)
    
    # Calculate the total price for each item
    df['TOTAL PRICE'] = df['ITEM QUANTITY'] * df['ITEM PRICE']
    
    # Group the data by ORDER ID
    order_groups = df.groupby('ORDER ID')
    
    # Process each order group
    for order_id, order_df in order_groups:
        # Sort the items in the order by ITEM NUMBER
        order_df = order_df.sort_values(by='ITEM NUMBER')
        
        # Calculate the grand total for the order
        grand_total = order_df['TOTAL PRICE'].sum()
        
        # Create a row for the grand total
        grand_total_row = pd.DataFrame([{col: '' for col in order_df.columns}])
        grand_total_row.iloc[0, order_df.columns.get_loc('ITEM NUMBER')] = 'GRAND TOTAL'
        grand_total_row.iloc[0, order_df.columns.get_loc('TOTAL PRICE')] = grand_total
        
        # Append the grand total row to the order DataFrame
        order_df = pd.concat([order_df, grand_total_row])
        
        # Define the output file path
        order_file = os.path.join(orders_dir, f"Order_{order_id}.xlsx")
        
        # Write the order DataFrame to an Excel file
        with pd.ExcelWriter(order_file, engine='xlsxwriter') as writer:
            order_df.to_excel(writer, index=False, sheet_name=f"Order {order_id}")
            
            # Get the workbook and worksheet objects
            workbook = writer.book
            worksheet = writer.sheets[f"Order {order_id}"]
            
            # Define formats for money and headers
            money_format = workbook.add_format({'num_format': '$#,##0.00'})
            header_format = workbook.add_format({'bold': True, 'align': 'center'})
            
            # Set column widths and formats
            column_settings = [
                ('A:A', 11),  # ORDER DATE
                ('B:B', 13),  # ITEM NUMBER
                ('C:C', 15),  # PRODUCT LINE
                ('D:D', 15),  # PRODUCT CODE
                ('E:E', 15),  # ITEM QUANTITY
                ('F:F', 13),  # ITEM PRICE
                ('G:G', 13),  # TOTAL PRICE
                ('H:H', 10),  # STATUS
                ('I:I', 30),  # CUSTOMER NAME
            ]
            
            for col, width in column_settings:
                worksheet.set_column(col, width)
            
            # Apply money format to ITEM PRICE and TOTAL PRICE columns
            price_cols = ['F:F', 'G:G']
            for col in price_cols:
                worksheet.set_column(col, 13, money_format)
            
            # Format the header row
            for col_num, value in enumerate(order_df.columns.values):
                worksheet.write(0, col_num, value, header_format)

if __name__ == '__main__':
    main()
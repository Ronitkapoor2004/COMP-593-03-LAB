import sys 
import os
from datetime import date
import pandas as pd

def main():
    sales_csv = get_sales_csv()
    orders_dir = create_orders_dir(sales_csv)
    process_sales_data(sales_csv, orders_dir)

# Get path of sales data CSV file from the command line
def get_sales_csv():
    # Check whether command line parameter provided
    # Check whether provide parameter is valid path of file
    if len(sys.argv) < 2:
        print("The path of sales data CSV file is not provided!!")
        sys.exit(1)
    else: 
        if os.path.exists(sys.argv[1]) == False:
            print("File not existing!!")
            sys.exit(2)
    return sys.argv[1]

# Create the directory to hold the individual order Excel sheets
def create_orders_dir(sales_csv):
    # Get directory in which sales data CSV file resides
    directory_path = os.path.dirname(os.path.abspath(sales_csv))
    # Determine the name and path of the directory to hold the order data files
    todays_dt = date.today().isoformat()
    order_directory_nm = f'orders_{todays_dt}'
    order_directory_path = os.path.join(directory_path,order_directory_nm)
    # Create the order directory if it does not already exist
    if os.path.exists(order_directory_path) == False:
        os.makedirs(order_directory_path)
    return order_directory_path

# Split the sales data into individual orders and save to Excel sheets
def process_sales_data(sales_csv, orders_dir):
    # Import the sales data from the CSV file into a DataFrame
    data = pd.read_csv(sales_csv)
    #print(data)
    # Insert a new "TOTAL PRICE" column into the DataFrame
    List1 = data['ITEM PRICE']
    List2 = list(List1)
    def multiply_itmqty_itmprice_func(item_qty):
        for pricevalue in List2:
            totalpriceearned = pricevalue*item_qty
            List2.remove(pricevalue)
            return totalpriceearned
    data.insert(7,'TOTAL PRICE',[multiply_itmqty_itmprice_func(item_qty) for item_qty in data['ITEM QUANTITY']] )
    # Remove columns from the DataFrame that are not needed
    data.drop(['ADDRESS','CITY','STATE','POSTAL CODE','COUNTRY'],axis=1,inplace=True)
    # Group the rows in the DataFrame by order ID
    group_of_data_acording_to_order_id = data.groupby('ORDER ID')
    # For each order ID:
    for orderid, data2 in group_of_data_acording_to_order_id: 
        # Remove the "ORDER ID" column
        data2.drop(['ORDER ID'],axis=1,inplace=True)
        # Sort the items by item number
        data2.sort_values(by = 'ITEM NUMBER',inplace = True)
        # Append a "GRAND TOTAL" row
        grandtotalvalue = f"${sum(data2['TOTAL PRICE'])}"
        new_row = {'ITEM PRICE':'GRAND TOTAL:','TOTAL PRICE': grandtotalvalue}
        data2.loc[len(data2)] = new_row
        # Determine the file name and full path of the Excel sheet
        file_path_of_excelsheet = os.path.abspath(orders_dir)
        file_name_contains_excelsheet = f"{orderid}.xlsx"
        absolute_file_path_of_excelsheet = os.path.join(file_path_of_excelsheet,file_name_contains_excelsheet)
        # Export the data to an Excel sheet
        data2.to_excel( absolute_file_path_of_excelsheet,index=False, sheet_name='Salesdata_orderidwise')

        # Format the Excel sheet 
        # Define format for the money columns
        # Format each colunm
        # close the sheet

        excel_writer = pd.ExcelWriter(absolute_file_path_of_excelsheet, engine="xlsxwriter")
        data2.to_excel(excel_writer,index=False, sheet_name='Salesdata_orderidwise')
        workbook = excel_writer.book
        excelworksheet = excel_writer.sheets['Salesdata_orderidwise']

        '''with pd.ExcelWriter(absolute_file_path_of_excelsheet) as writer:
        data.to_excel(writer)
        workbook = writer.book
        worksheet = writer.sheets['Salesdata_orderidwise']'''
        
        formatofcolumn = workbook.add_format({"num_format": "$#,##0.000"})
        excelworksheet.set_column('A:A',11)
        excelworksheet.set_column('B:B',13)
        excelworksheet.set_column('C:C',15)
        excelworksheet.set_column('D:D',15)
        excelworksheet.set_column('E:E',15)
        excelworksheet.set_column('F:F',13,formatofcolumn)
        excelworksheet.set_column('G:G',13,formatofcolumn)
        excelworksheet.set_column('H:H',10)
        excelworksheet.set_column('I:I',30)
        excel_writer.close()
        
    

if __name__ == '__main__':
    main()
# Lucas Miller
# August 12th, 2022
# Created for FortisureIT's pre-employment aptitude screening project.

import pandas as pd
from datetime import date
import queue

action_log = queue.Queue()

def remove_duplicates(sales_order_header):
    # Removes the duplicate entries from SalesOrderHeader
    action_log.put("remove_duplicates(sheet1)")
    print("# Duplicates removed from SalesOrderHeader")
    sales_order_header = sales_order_header.drop_duplicates('SalesOrderID', keep='first')
    return sales_order_header

def update_unit_price(product, sales_order_detail):
    # Fixes Unit Price for specified productID
    action_log.put("update_unit_price(" + product + ", sheet1)")
    sales_order_detail.loc[sales_order_detail.eval('ProductID==' + product), 'UnitPrice'] = sales_order_detail.loc[sales_order_detail.eval('ProductID==' + product_id), 'UnitPrice'] / 2.0
    return sales_order_detail

def round_dollar_amounts(sales_order_detail, sales_order_header):
    # Rounds all dollar amounts to 2 decimals
    action_log.put("round_dollar_amounts(sheet1, " + "sheet2)")
    print("# All dollar amounts rounded successfully\n")
    sales_order_detail = sales_order_detail.round(decimals = 2)
    sales_order_header = sales_order_header.round(decimals = 2)
    return sales_order_detail

def verify_line_total(sales_order_detail):
    # Verify that UnitPrice * OrderQty is equal to LineTotal
    action_log.put("verify_line_total(sheet1)")
    # On Sales Order Detail, the Unit Price multiplied by the Order Qty should equal the Line Total.
    
    values = ((sales_order_detail['UnitPrice'].astype(float) - (sales_order_detail['UnitPrice'].astype(float) 
    * sales_order_detail['UnitPriceDiscount'].astype(float))) * sales_order_detail['OrderQty'].astype(float))
    
    # Rounding for accurate precision
    values = values.round(decimals = 8)
    #sales_order_detail = sales_order_detail.round(decimals = 8)

    print("# Validating Updated Unit Price...\n")
    
    if(values.equals(sales_order_detail["LineTotal"])):
        print("# LineTotal Validation Successful\n")
    else:
        print("# Validation of LineTotal unsuccessful, Printing Differences")
        print(values.compare(sales_order_detail['LineTotal']), "\n")
    return sales_order_detail

def verify_data(sales_order_detail, sales_order_header):
    # Verify that SalesOrderHeader SubTotal sum is equivalent to SalesOrderDetail LineTotal sum 
    action_log.put("verify_data(sheet1, sheet2)")
    sub_total_sum = sheet2['SubTotal'].sum()
    line_total_sum = sheet1['LineTotal'].sum()
    sub_total_sum = sub_total_sum.round(decimals=0)
    line_total_sum = line_total_sum.round(decimals=0)

    print("# SalesOrderHeader SubTotal:  ", sub_total_sum)
    print("# SalesOrderDetail LineTotal: ", line_total_sum, "\n")

    if(line_total_sum == sub_total_sum):
        print("# LineTotal/Subtotal Validated Successfully:")
        print("# LineTotal: ", line_total_sum, "==", sub_total_sum, ": SubTotal\n")
    else:
        print("# Totals are not equivalent, Validation Failed\n")

    return sales_order_detail

def write_to_file(name, cur_date, ws1, ws2, ws3, ws4, ws5, ws6):
    # Write changes to new file, append name and date to filename 
    with pd.ExcelWriter((name + " - FIT Sales Data " + cur_date + ".xlsx"), engine="xlsxwriter") as writer:
        ws1.to_excel(writer, sheet_name="Sales Order Detail", index=False)
        ws2.to_excel(writer, sheet_name="Sales Order Header", index=False)
        ws3.to_excel(writer, sheet_name="Sales Reason", index=False)
        ws4.to_excel(writer, sheet_name="Sales Order Header w Reason", index=False)
        ws5.to_excel(writer, sheet_name="Sales Territory", index=False)
        ws6.to_excel(writer, sheet_name="Action and Parameter List", index=False)



excel_workbook = 'FortisureIT Pre-Employment Sales Data - Developer.xlsx'
sheet1 = pd.read_excel(excel_workbook, sheet_name='Sales Order Detail')
sheet2 = pd.read_excel(excel_workbook, sheet_name='Sales Order Header')
sheet3 = pd.read_excel(excel_workbook, sheet_name='Sales Reason')
sheet4 = pd.read_excel(excel_workbook, sheet_name='Sales Order Header w Reason')
sheet5 = pd.read_excel(excel_workbook, sheet_name='Sales Territory')

print("# 1. Remove Duplicates\n")
print("# 2. Update Unit Price\n")
print("# 3. Round All Dollar Amounts\n")
print("# 4. Update Last Modified Date\n")
process_selection = input("# What processes would you like to run?: ")
action_log.put("Program Entry Point: " + str(process_selection))


if("1" in process_selection):
    # Removes the duplicate SalesOrderID's from SalesOrderHeader
    sheet2 = remove_duplicates(sheet2)
    
if("2" in process_selection):
    product_id = input("# Enter ProductID To Be Modified: ")
    action_log.put("Enter ProductID: " + product_id) 
    sheet1 = update_unit_price(product_id, sheet1)
    sheet1 = verify_line_total(sheet1)     

    verify_data(sheet1, sheet2)

if("3" in process_selection):
    # Rounds all dollar amounts in the dataframe to the nearest hundredths place
    sheet1 = round_dollar_amounts(sheet1, sheet2)

print("# Final Results")
print(sheet1.head(20), "\n")

user_name = input("Enter your name: ")
now = date.today()
current_date = now.strftime("%m-%d-%Y")

action_log.put("write_to_file( " + name + ", " + cur_date + ", " + "sheet1, " + "sheet2, " + "sheet3, " + "sheet4, " + "sheet5, " + "df")
a = list(action_log.queue)
df = pd.DataFrame(a)
df.columns = ["Actions"]
write_to_file(user_name, current_date, sheet1, sheet2, sheet3, sheet4, sheet5, df)
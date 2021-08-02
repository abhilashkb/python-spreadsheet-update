import openpyxl
from openpyxl import cell

orderfile=openpyxl.load_workbook("orders.xlsx")


customer_loc=orderfile["Sheet1"]
customer_phone=orderfile["Sheet2"]
customer_orders=orderfile["Sheet3"]

for customer in range(2,customer_loc.max_row+1):
    print(customer_loc.cell(customer,2).value)
    id=int(customer_loc.cell(customer,1).value) 
    for phone in range(2,customer_phone.max_row+1):
        if ( id == int(customer_phone.cell(phone,1).value)) :
            customer_loc.cell(customer,5).value=int(customer_phone.cell(phone,2).value)
            
            
orderfile.save("solution.xlsx")
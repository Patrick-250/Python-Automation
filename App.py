import openpyxl as xl
wb=xl.load_workbook("transactions.xlsx")
sheet=wb["Sheet1"]
#cell=sheet1["a1"]
#print(cell)
cell=sheet.cell(1,1)
#print(cell.value)
number_of_rows=sheet.max_row
#print(number_of_rows)
for row in range(2,number_of_rows+1): #loop start from 2 to ignore the headers
    print(row)










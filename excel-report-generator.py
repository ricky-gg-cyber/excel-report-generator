import csv
import sys
import argparse
import openpyxl
from openpyxl.styles import Font

def xlxs_writing_loop(dictionary:dict,sheet,interested_column):
    row_count=1
    for categories,quantity in dictionary.items():
        row_count+=1
        sheet.cell(row=row_count,column=interested_column).value=categories
        categories_quantity_cell=sheet.cell(row=row_count,column=interested_column+1)
        categories_quantity_cell.number_format='0.00'
        categories_quantity_cell.value=quantity

try:
    parser=argparse.ArgumentParser()
    parser.add_argument('file_csv',help='The csv file to analyze')
    parser.add_argument('output_xlsx', help='The xlsx output file')
    args=parser.parse_args()
    filename=args.file_csv
    with open(filename,newline='',encoding='utf-8-sig') as csvfile:
        reader=csv.DictReader(csvfile)
        data=list(reader)
except FileNotFoundError:
    print('Error,file not found')
    sys.exit(1)
total=0
products_dict=dict()
categories_dict=dict()
row_number=1
try:
    for i in data:
        row_number+=1
        if None in i:
            print(f"Error on line {row_number}: too many columns")
            sys.exit(1)
        products_dict.setdefault(i['product'],0)
        categories_dict.setdefault(i['category'],0)
        raw_total=int(i['quantity'])*float(i['unit_price'].replace(',','.'))
        products_dict[i['product']]+=raw_total
        categories_dict[i['category']]+=raw_total
        total+=raw_total
except KeyError:
    print('One column is missing from the expected input')
    sys.exit(1)
except ValueError:
    print(f'invalid quantity or price found at row: {row_number}')
    sys.exit(1)
wb=openpyxl.Workbook()
sheet1=wb.active
sheet1.title='Data'
sheet2=wb.create_sheet(title='Recap')
sheet1.cell(row=1,column=1,value='Products').font=Font(bold=True)
sheet1.cell(row=1,column=2,value='Products quantity').font=Font(bold=True)
sheet1.cell(row=1,column=4,value='Categories').font=Font(bold=True)
sheet1.cell(row=1,column=5,value='Categories quantity').font=Font(bold=True)
xlxs_writing_loop(products_dict,sheet1,1)
xlxs_writing_loop(categories_dict,sheet1,4)
last_row=sheet1.max_row
sheet1.cell(row=last_row+1,column=1,value='Total').font=Font(bold=True)
total_quantity_cell=sheet1.cell(row=last_row+2, column=2)
total_quantity_cell.number_format='0.00'
total_quantity_cell.value=total
wb.save(args.output_xlsx)
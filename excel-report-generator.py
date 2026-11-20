import csv
import sys
import argparse
import openpyxl
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

REQUIRED_COLUMNS = {'product', 'category', 'quantity', 'unit_price'}


def read_csv(path):
    try:
        with open(path, newline='', encoding='utf-8-sig') as csvfile:
            reader = csv.DictReader(csvfile)
            rows = list(reader)
            return reader.fieldnames or [], rows
    except FileNotFoundError:
        print('Error: file not found')
        sys.exit(1)
    except OSError as error:
        print(f'Error: cannot read the file ({error.strerror})')
        sys.exit(1)


def write_row(sheet, values, row, bold=False):
    for column, value in enumerate(values, start=1):
        sheet.cell(row=row, column=column, value=value).font = Font(bold=bold)


def write_totals(sheet, totals, column, title):
    sheet.cell(row=1, column=column, value=title).font = Font(bold=True)
    sheet.cell(row=1, column=column + 1, value='Total Sales').font = Font(bold=True)
    for row, (name, amount) in enumerate(totals.items(), start=2):
        sheet.cell(row=row, column=column, value=name)
        cell = sheet.cell(row=row, column=column + 1, value=round(amount, 2))
        cell.number_format = '0.00'


def adjust_column_widths(workbook):
    for sheet in workbook.worksheets:
        for column in sheet.columns:
            longest = max((len(str(cell.value)) for cell in column
                           if cell.value is not None), default=0)
            letter = get_column_letter(column[0].column)
            sheet.column_dimensions[letter].width = max(longest + 3, 10)


def main():
    parser = argparse.ArgumentParser(
        description='Create an Excel report from a CSV file of sales.')
    parser.add_argument('file_csv', help='The csv file to analyze')
    parser.add_argument('output_xlsx', help='The xlsx output file')
    args = parser.parse_args()

    fieldnames, data = read_csv(args.file_csv)
    missing = REQUIRED_COLUMNS - set(fieldnames)
    if missing:
        print(f"Error: missing columns: {', '.join(sorted(missing))}")
        sys.exit(1)
    if not data:
        print('Error: no data rows found in the file')
        sys.exit(1)

    wb = openpyxl.Workbook()
    sheet1 = wb.active
    sheet1.title = 'Data'
    sheet2 = wb.create_sheet(title='Recap')
    write_row(sheet1, fieldnames, 1, bold=True)
    price_column = fieldnames.index('unit_price') + 1

    total = 0
    products_dict = dict()
    categories_dict = dict()
    row_number = 1
    for row in data:
        row_number += 1
        if None in row:
            print(f'Error on line {row_number}: too many columns')
            sys.exit(1)
        if None in row.values():
            print(f'Error on line {row_number}: missing values')
            sys.exit(1)
        try:
            quantity = int(row['quantity'])
            price = float(row['unit_price'].replace(',', '.'))
        except ValueError:
            print(f'Error on line {row_number}: invalid quantity or price')
            sys.exit(1)

        row_values = []
        for name in fieldnames:
            if name == 'quantity':
                row_values.append(quantity)
            elif name == 'unit_price':
                row_values.append(price)
            else:
                row_values.append(row[name])
        write_row(sheet1, row_values, row_number)
        sheet1.cell(row=row_number, column=price_column).number_format = '0.00'

        products_dict.setdefault(row['product'], 0)
        categories_dict.setdefault(row['category'], 0)
        row_total = quantity * price
        products_dict[row['product']] += row_total
        categories_dict[row['category']] += row_total
        total += row_total

    write_totals(sheet2, products_dict, 1, 'Products')
    write_totals(sheet2, categories_dict, 4, 'Categories')
    total_row = sheet2.max_row + 2
    sheet2.cell(row=total_row, column=1, value='Total').font = Font(bold=True)
    total_cell = sheet2.cell(row=total_row, column=2, value=round(total, 2))
    total_cell.number_format = '0.00'
    total_cell.font = Font(bold=True)

    adjust_column_widths(wb)
    try:
        wb.save(args.output_xlsx)
    except (FileNotFoundError, PermissionError):
        print('Error: cannot save the file. Check that the folder exists '
              'and that the file is not open in another program.')
        sys.exit(1)
    print(f'Report saved to {args.output_xlsx}')


if __name__ == '__main__':
    main()
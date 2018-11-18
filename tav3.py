
import csv
import re

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter


DEST_FILENAME = '/home/anonymous/cashbookTaxYr2018-2019.xlsx'
SPACE_AND_CHECK_COL = 2
SPACE_AND_TOTAL_BOX = 2
MIN_TYPES_TRANSACTION = 11

cashbook = load_workbook(DEST_FILENAME)
print("Cashbook opened")

receipts = cashbook['Cashbook Receipts']
payments = cashbook['Cashbook Payments']
dla = cashbook["Director's Loan Account"]

#print(get_column_letter(11) + " " + get_column_letter(receipts.max_column - SPACE_AND_CHECK_COL))
#print(get_column_letter(11) + " " + get_column_letter(payments.max_column - SPACE_AND_CHECK_COL))

transaction_type_dict = {}
check_desc = []

# add catagories from file
with open('/home/anonymous/Dropbox/repeated_transactions.csv', newline='') as csvfile:
	reader = csv.reader(csvfile)
	for row in reader:
		transaction_type_dict[row[0]] = row[1]
		check_desc.append(''.join(row[0]))

hold = ""

for i in check_desc:
	hold = hold + i + '|'
hold = hold[:-1]  # remove last '|'

first_two = re.compile(hold)

for row in range(6,receipts.max_row):		# Description Column
	if receipts.cell(column=3, row=row).value != None:
		mo = first_two.search(receipts.cell(column=3, row=row).value)
		if mo != None:
			receipts.cell(column=9, row=row, value=transaction_type_dict[mo.group()])


for row in range(6,payments.max_row):		# Description Column
	if payments.cell(column=3, row=row).value != None:
		mo = first_two.search(payments.cell(column=3, row=row).value)
		if mo != None:
			payments.cell(column=9, row=row, value=transaction_type_dict[mo.group()])

print(dla.max_row)

for row in receipts.iter_cols(min_row=6, min_col=9, max_col=9 , max_row=receipts.max_row-SPACE_AND_TOTAL_BOX):
	for cell in row:
		if cell.value == "Director’s Loan Account":
			print(receipts['B' + str(cell.row)].value)
			print(receipts['C' + str(cell.row)].value)
			print(receipts['D' + str(cell.row)].value)

for row in payments.iter_cols(min_row=6, min_col=9, max_col=9, max_row=payments.max_row-SPACE_AND_TOTAL_BOX):
	for cell in row:
		if cell.value == "Director’s Loan Account":
			print('B' + str(cell.row))
			print('C' + str(cell.row))
			print('D' + str(cell.row))

print(dla['A3'].value)

cashbook.save(DEST_FILENAME)
print("Cashbook closed")



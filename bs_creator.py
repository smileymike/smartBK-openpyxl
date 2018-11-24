import glob
from openpyxl import load_workbook

CASHBOOKS_FOLDER_LOCATION = '/home/anonymous/Cashbooks'
LAST_TAXYEAR = CASHBOOKS_FOLDER_LOCATION + '/cashbookTaxYr2017-2018.xlsx'
CURRENT_TAXYEAR = CASHBOOKS_FOLDER_LOCATION + '/cashbookTaxYr2018-2019.xlsx'

files = []

for name in glob.glob(CASHBOOKS_FOLDER_LOCATION + '/cashbookTaxYr20[1-9][0-9]-20[1-9][0-9].xlsx'):
	files.append(name)

print('Start Test')

for each_file in files:
	try:
		cashbook = load_workbook(each_file)
		print(each_file +" Cashbook opened")

#		receipts = cashbook['Cashbook Receipts']
#		payments = cashbook['Cashbook Payments']
#		dla = cashbook["Director's Loan Account"]
#		pla = cashbook["Profit & Loss Account"]
		bs = cashbook["Balance Sheet"]
	except:
		print("Error in opening a cashbook")

print('End Test')

cashbook = load_workbook(LAST_TAXYEAR)
print('Cashbook opened again!')

bs = cashbook["Balance Sheet"]
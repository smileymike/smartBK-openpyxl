import csv

transactions = []

# Read data.csv from Barclay Bank Website (download first)
with open('data.csv', newline='') as csvfile:
	readCSV = csv.reader(csvfile, delimiter=',')
	for row in readCSV:
		transactions.append(row)

for each_transaction in reversed(transactions):
	if each_transaction[3] != 'Amount':
		print(each_transaction)
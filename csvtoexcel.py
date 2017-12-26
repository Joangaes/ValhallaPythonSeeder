import os
import glob
import csv
from xlsxwriter.workbook import Workbook


for csvfile in glob.glob(os.path.join("E:\Jose\Documents\Mutuo\MigrateDBValhalla\MigrateDBValhalla/Uber_To_Arkafin/*.csv")):
    workbook = Workbook(csvfile[:-4] + '.xlsx')
    worksheet = workbook.add_worksheet()
    with open(csvfile, 'rt') as f:
        reader = csv.reader(f)
        for r, row in enumerate(reader):
            for c, col in enumerate(row):
                worksheet.write(r, c, col)
    workbook.close()

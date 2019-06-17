from openpyxl import Workbook, load_workbook
import xml.etree.ElementTree as ET

wb = load_workbook('master_data.xlsx')

for sheet in wb.sheetnames:
    print(sheet)

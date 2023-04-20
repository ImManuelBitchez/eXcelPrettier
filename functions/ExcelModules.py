from openpyxl import Workbook,load_workbook
import numpy as np
from pathlib import Path     
projectFolder = Path("C:/Users/manue/Documents/Stellantis/Python/AWS/")

def formatName(filename):
    temp = filename.split('\\')
    rawName = temp[len(temp)-1].split('.')
    name = rawName[0]
    return name


def transpose(workbook,fileName):
    wb = load_workbook(workbook)
    ws = wb.active
    max_row = ws.max_row
    max_col = ws.max_column
    # Selezionare l'intervallo di celle da trasporre
    cell_range = ws.calculate_dimension()
    range_to_transpose = ws[cell_range]
    transposed_range = np.transpose([[cell.value for cell in row] for row in range_to_transpose])
    for row in transposed_range:
        ws.append(list(row))
    ws.delete_rows(1,max_row)
    wb.save(fileName)


def createFile(workbook):
    outfile = Workbook()
    outWS = outfile.active
    wb = load_workbook(workbook)
    ws = wb.active
    cell_range = ws.calculate_dimension()
    range_to_copy = ws[cell_range]
    for row in cell_range:
        cell_list = []
        for cell in row:
            cell_list.append(cell.value)
    print(cell_list)
    #outWS.append(cell_list)
    #outfile.save("out.xlsx")
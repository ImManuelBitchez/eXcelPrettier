import sys,os
import win32com.client

from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from pathlib import Path
from functions.ExcelModules import *

projectFolder = Path("C:/Users/manue/Documents/Stellantis/Python/AWS/")
module_path = Path(f"{projectFolder}\modules\Modulo1.bas")

if(1==2): #len(sys.argv) <= 1
    print("No csv file founded!")
    print("Please insert at least one file")
    os._exit(1)
else:
    try:
        baseFile = sys.argv[1]
        fileName = formatName(sys.argv[1])
        outfile = f"{projectFolder}\\files\\{fileName}.xlsx"

        excel = win32com.client.Dispatch('Excel.Application')
        wb = excel.Workbooks.Open(baseFile)
        wb.VBProject.VBComponents.Import(module_path)
        excel.Application.Run("FormatCSV")
        wb.SaveAs(outfile,51)
        excel.Quit()
        print("Macro ran succesfully!")

    except Exception as e:
        print(e)
    
 
    transpose(workbook=outfile,fileName=outfile)
    copyToFile(outfile,filename=outfile)
  
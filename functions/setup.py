import os,platform
from pathlib import Path
from functions.ExcelModules import *


def findWinUser():
    paths = os.environ['PATH'].split(";")
    for path in paths:
        temp = path.split("\\")
        for char in temp:
            if(char == "Users"):
                user = temp[temp.index(char)+1]
                return user
        
def initProjectFolder(baseFolder):
    dirs = os.listdir(baseFolder)
    if "Consumption" in dirs:
        print("Project Folder already exists!")
        return 0
    else:
        os.mkdir(os.path.join(baseFolder,"Consumption"))
        os.mkdir(os.path.join(baseFolder,"Consumption","Excel"))
        
def init():
    p = platform.system()
    if(p == 'Linux'):
        pass
    elif(p == 'Darwin'):
        pass
    elif(p == "Windows"):
        user = findWinUser()
        baseFolder = Path(f"C:\\Users\\{user}\\Documents\\")
        initProjectFolder(baseFolder)
        projectFolder = os.path.join(baseFolder,"Consumption")
        outFile = os.path.join(projectFolder,"consumption.xlsx")
        excelFolder = os.path.join(projectFolder,"Excel\\")
        return {
            "user": user,
            "projectFolder": projectFolder,
            "outFile": outFile,
            "excelFolder": excelFolder 
        }




# import openpyxl
import pandas
import os
import re
import linecache
import shutil

# Please put all xlsx or xls files in the directory, and make sure all datas are in the Sheet1(default)
# Note: Can't process multiple Sheets within a single workbook
def main():
    print("Begin to Process:")
    
    # dirExcel=os.chdir("./excelFile/")
    dirList=os.listdir("./excelFile")
    print(dirList)
    i=0
    for l in dirList:
        # if re.search("^.*xlsx$",l) or re.search("^.*xls$",l):
        # Put all excel files into dir:excelFile
            print(l)
            x='./excelFile/'+l
            print(x)
            xls = pandas.read_excel(x, 'Sheet1', index_col=None)
            xls.to_csv('toCsv.csv', encoding='utf-8')
            # shutil.move('./toCsv.csv','./csv/toCsv.csv')
            os.rename("./toCsv.csv","csv"+str(i)+".csv")
            i+=1

    
    finalCsv=open("./final.csv",mode='a')
    with open('./final.csv','w') as fi:
        fi.write(',col1,col2,...\n')
    for ll in os.listdir("./"):
        if re.search(r"^csv[0-9]*\.csv$",ll):
            # Change the number or write other codes based the certain suitation
            finalCsv.write(linecache.getline(ll,2))
            y='./'+ll
            print(y)
            os.remove(y)
    
    # Csv is ok, don't have to convert to xlsx   
    # But why didn't it work? 
    fCsv=pandas.read_csv('./final.csv').to_excel('./finall.xlsx','Sheet1')
    # fCsv.to_excel('./final.xlsx')
    

    finalCsv.close()


if __name__=="__main__":
    main()

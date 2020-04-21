# coding: utf-8

import pandas
import os
import json
import sys

inputDir = 'example'
tmpDir = 'tmp'
resultFile = 'Result.xlsx'
SkipRowMap = {'02':1, '03':3, '04':3}

LogFile = 'log.txt'
f = open(LogFile, 'w')
old = sys.stdout
sys.stdout = f

SheetCols = {}
HeaderChecker = []

def cleanCache():
    if os.path.exists(tmpDir):
        for filename in os.listdir(tmpDir):
            os.remove('{0}/{1}'.format(tmpDir,filename))
        os.removedirs(tmpDir)

def doExcel(xlsFilename):
    print('DoExcel:', xlsFilename)
    pf = pandas.ExcelFile(xlsFilename)
    for sheet in pf.sheet_names:
        sheetNum = sheet.split(' ')[0]
        if sheetNum in SkipRowMap:
            pSheet = pandas.read_excel(pf, sheet,skiprows=SkipRowMap[sheetNum], index_col=0)
            isHeader = False
            if sheet not in HeaderChecker:
                isHeader = True
                HeaderChecker.append(sheet)
            #print('ReadExcel File：{}，Sheet:{}，列数：{}'.format(xlsFilename, sheet, pSheet.shape[1]))
            if sheet not in SheetCols:
                SheetCols[sheet] = {}
            SheetCols[sheet][xlsFilename] = pSheet.shape[1]
            pSheet.to_csv('{0}/{1}.csv'.format(tmpDir,sheet), mode='a+', header=isHeader)

def sumAll():
    with pandas.ExcelWriter(resultFile) as writer:
        for filename in os.listdir(tmpDir):
            sheetName = filename.split('.')[0]
            print('SumSheet:',sheetName)
            pCsv = pandas.read_csv('{0}/{1}'.format(tmpDir,filename))
            pCsv.to_excel(writer, sheet_name=sheetName, header=False)
        

cleanCache()
os.mkdir(tmpDir)

for filename in os.listdir(inputDir):
    doExcel('{0}/{1}'.format(inputDir,filename))

print('\nSheetCheck:{}\n'.format(json.dumps(SheetCols, indent=4, ensure_ascii=False)))
sumAll()

cleanCache()

print('\nSuccessFinish！！！')
sys.stdout = old
f.close()

print('输出文件：{0}'.format(resultFile))
os.system('pause')
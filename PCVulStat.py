# -*- coding: utf-8 -*-

import os
import json
import openpyxl

def search(dirname):
    """이 함수는 디렉토리(폴더) 경로를 받아 해당 디렉토리의 파일 리스트를 반환함.

    : param: str dirname: 디렉토리(폴더) 경로

    예제:
        다음과 같이 사용:
        >>> filelist = search("C:/temp")
            print(filelist)
            ['C:/temp/test1', 'C:/temp/test2', 'C:/temp/test3']
 
    :returns list: 파일경로 리스트 반환
    """

    fullFilenames = []
    filenames = os.listdir(dirname)
    for filename in filenames:
        fullFilename = os.path.join(dirname, filename)
        fullFilenames.append(fullFilename)

    # print (fullFilenames)

    return fullFilenames

def jsonRead(filename):
    """이 함수는 json 파일을 읽어 python dictionary 값으로 반환.

    : param: str filename: 파일 경로.

    예제:
        다음과 같이 사용:
        >>> secuCheckList = jsonRead("SecuCheckList.json")

    :returns dict: json 파일 내용을 dictionary 타입 값으로 반환
    """
    fileData = open(filename, "r").read()
    jsonDataAsPythonValue = json.loads(fileData)
    return jsonDataAsPythonValue

# 결과 파일 경로 리스트
fullFilenames = search("201805")

#print(fullFilenames)
#print(type(fullFilenames))

# 점검 항목 리스트
secuCheckDict = jsonRead("SecuCheckList.json")
secuCheckList = list(secuCheckDict.items())

#print(secuCheckDict)
#print(type(secuCheckDict))
#print(len(secuCheckList))

# 엑셀 파일 초기화
wb = openpyxl.Workbook()
ws = wb.active
ws.cell(row=1, column=1, value=u"코드")
ws.cell(row=1, column=2, value=u"항목명")
for row in range(2, len(secuCheckList)+2):
    for col in range(1, 3):
        ws.cell(row=row, column=col, value=secuCheckList[row-2][col-1])
wb.save("/Users/luminary_topco/Desktop/201805-Topco_PC_Security_Check.xlsx")

# 엑셀 파일에 점검 결과 입력
wb = openpyxl.load_workbook("/Users/luminary_topco/Desktop/201805-Topco_PC_Security_Check.xlsx")
ws = wb.active
colNum = 3
for fullFilename in fullFilenames:
    secuCheckData = jsonRead(fullFilename)
    ws.cell(row=1, column=colNum, value=secuCheckData['name'])
    for row in range(2, len(secuCheckList)+2):
        ws.cell(row=row, column=colNum, value=secuCheckData[secuCheckList[row-2][0]])

        #print(secuCheckData[secuCheckList[row-2][0]])

    colNum += 1

    #print(colNum)

wb.save("/Users/luminary_topco/Desktop/201805-Topco_PC_Security_Check.xlsx")

#    print(secuCheckData)
#    print(type(secuCheckData))



# -*- coding: utf-8 -*-

import os
import json
import openpyxl

def search(dirname):
    fullFilenames = []
    filenames = os.listdir(dirname)
    for filename in filenames:
        fullFilename = os.path.join(dirname, filename)
        fullFilenames.append(fullFilename)

    print (fullFilenames)

    return fullFilenames

def jsonRead(filename):
    fileData = open(filename, "r").read()
    jsonDataAsPythonValue = json.loads(fileData)
    return jsonDataAsPythonValue

fullFilenames = search("201805")

print(fullFilenames)
print(type(fullFilenames))

secuCheckList = jsonRead("SecuCheckList.json")

print(secuCheckList)
print(type(secuCheckList))

for fullFilename in fullFilenames:
    secuCheckData = jsonRead(fullFilename)

    print(secuCheckData)
    print(type(secuCheckData))




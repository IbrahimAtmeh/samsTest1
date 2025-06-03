import time
import cv2
import numpy
import pytesseract
import zxingcpp
import requests
from openpyxl import load_workbook
import win32com.client as win32
import streamlit as st

companyName = "empty"
productName = "empty"
expDate = "empty"
pkgQty = "empty"
dimensions = "empty"
numDevices = "empty"
results,dataMatrixCode, rawCode = "","",""

tInit = time.time()
tElapsed = 0
tryCode = True

def get_company_name(device_id):
    #what we want to find
    global dimensions, numDevices, companyName,productName, pkgQty
    # The API endpoint (website address for the service)
    url = "https://accessgudid.nlm.nih.gov/api/v3/devices/lookup.json"
    
    # Parameters (the information we're sending)
    params = {
        "di": device_id,  # di stands for "Device Identifier"
        "fields": 'companyName,deviceDescription,deviceCount'
    }
    
    # Sending the request
    start = time.time()
    response = requests.get(url, params=params)
    timeNeeded = time.time()-start
    print (timeNeeded)
    # Checking if the request was successful
    if response.status_code == 200:
        data = response.json()  # Convert the response to Python dictionary
        #print(data)
        # Navigating through the response to find company name
        if 'gudid' in data and 'device' in data['gudid']:
            global dimensions, numDevices, companyName,productName, pkgQty
            dimensions = data['gudid']['device'].get('deviceSizes','N/A')
            numDevices = data['gudid']['device'].get('deviceCount')
            pkgQty = data['gudid']['device']['identifiers']['identifier'][0].get('pkgQuantity')
            productName = data['gudid']['device'].get('deviceDescription')
            companyName = data['gudid']['device'].get('companyName')
            print (dimensions)
            print (numDevices)
            print (pkgQty)
            print (productName)
            print (companyName)

#load excel
excel = win32.Dispatch("Excel.Application")
wb=excel.Workbooks.Open(r"C:\samsProgram\secondTest.xlsx")
ws = wb.Worksheets("Sheet1")

url = "http://10.0.0.4:8080/video"  # IPv4 + /video endpoint
cap = cv2.VideoCapture(url)

while True:
    dataMatrixCode = ""
    rawCode = ""
    ret, frame = cap.read()
    if not ret:
        print("Failed to receive frame. Check IP/port!")
        break
    cv2.imshow('Phone Camera', frame)
    results = zxingcpp.read_barcodes(frame)
    if results !=[]:
       dataMatrixCode = results[0].text
    if len(dataMatrixCode)>14:
        rawCode = dataMatrixCode
        if time.time()-tInit > 2:
            tryCode = True
        if tryCode == True: 
            tInit = time.time()
            tryCode = False
            nextRowNum = 1
            while ws.Cells(nextRowNum,1).Value is not None:
                nextRowNum += 1           
            di = "empty"
            parts = rawCode.split('(')

            for p in parts:
                sub = p.split(')')
                if len(sub)>1:
                    signal,data = sub[0],sub[1]
                    if(signal =="01"):
                        di = data
                    if(signal =="17"):
                        expDate = data

            device_id = di  
            get_company_name(device_id)
            if dimensions == None:
                dimensions = "empty"
            if pkgQty == None:
                pkgQty = "empty"
            if numDevices == None:
                numDevices = "empty"
            if companyName == None:
                companyName = "empty"
            if productName == None:
                productName = "empty"

            year = expDate[:2]  # First 2 digits (27)
            month = expDate[2:4] # Next 2 digits (11)
            day = expDate[4:]
            expDate = f"{month}/{day}/{year}"

            print(expDate)

            #print values to excel and save
            ws.Cells(nextRowNum, 1).Value = companyName
            ws.Cells(nextRowNum, 2).Value = productName
            ws.Cells(nextRowNum, 3).Value = expDate
            ws.Cells(nextRowNum, 4).Value = dimensions
            ws.Cells(nextRowNum, 5).Value = pkgQty
            ws.Cells(nextRowNum, 6).Value = numDevices

            wb.Save()
        
        



    if cv2.waitKey(1) == ord('q'):
        break

cap.release()
cv2.destroyAllWindows()



excel.Quit()


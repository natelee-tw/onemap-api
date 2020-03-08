import json
import requests
import openpyxl
import time

#Open worksheet 
wb = openpyxl.load_workbook('List of HDB Address.xlsx')
sheet = wb['sheet1']

count = 0 

for row in range(2,sheet.max_row+1):
        if sheet['J' + str(row)].value is not None:
            print ("Existing, move to next line")
            continue
        elif count < 200: #retrive 
            query_address=sheet['A' + str(row)].value

#import request 
#formulate query string       
            query_string='https://developers.onemap.sg/commonapi/search?searchVal='+str(query_address)+'&returnGeom=Y&getAddrDetails=Y&pageNum=1'
            resp = requests.get(query_string)
        
#Convert JSON into Python Object 
            data=json.loads(resp.content)
        #print(type(data))

#Extract data from JSON
            try:
                sheet['B' + str(row)]=data['results'][0]['LONGITUDE']
                sheet['C' + str(row)]=data['results'][0]['LATITUDE']
                sheet['D' + str(row)]=data['results'][0]['SEARCHVAL']
                sheet['E' + str(row)]=data['results'][0]['BLK_NO']
                sheet['F' + str(row)]=data['results'][0]['ROAD_NAME']
                sheet['G' + str(row)]=data['results'][0]['BUILDING']
                sheet['H' + str(row)]=data['results'][0]['ADDRESS']
                sheet['I' + str(row)]=data['results'][0]['POSTAL']
                sheet['J' + str(row)]= 1
                print (str(query_address)+" Lat: "+data['results'][0]['LONGITUDE']+",Long: "+data['results'][0]['LATITUDE'])
            except:
                sheet['D' + str(row)]= 1
                print ("Error")
                pass
            
            count = count + 1   
        else: 
            print ("Pausing for 15 Seconds")
            wb.save (('all_zipcodes.xlsx'))
            time.sleep(15)
            count = 0

        wb.save('List of HDB Address.xlsx')
        
print('Done.')

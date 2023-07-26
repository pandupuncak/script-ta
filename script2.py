import json
import requests
import pygsheets
from openpyxl import load_workbook

client = pygsheets.authorize(service_file='sdi-script-4b8593c0b4ed.json')

# # test
# link = "http://geoportal.mubakab.go.id/geoserver/wms?service=WFS&version=1.0.0&request=GetFeature&typeName=Bappeda:kantordesa_pt_50k_160620220307023727&outputFormat=shape-zip"
# file = requests.get(link)
# print(file.status_code)
# if file.status_code == 200:
#     print("HEY")
# print(file)
# # jsonResponse = file.json()
# # print(jsonResponse)

spreadsheet = load_workbook(filename = 'Addition.xlsx')
worksheet = spreadsheet.get_sheet_by_name("Sheet1")

originalLink = "https://katalog.data.go.id/api/action/package_show?id="
APILink = "https://katalog.data.go.id/api/3/action/datastore_search?resource_id="

iteration = int(input("How many iterations? "))
current_row = int(input("Starting row: "))

def verify_metadata_api(metadata,i):
    for row in metadata["resources"]:
        try:
            if row["format"] == "WMS" or row["format"] == "WFS":
                if row["url"]:
                    urlTest = requests.get(row["url"], timeout=5)
                    if urlTest.status_code == 200:
                        return 1
            else:
                if row["package_id"]:
                    url = APILink + row["id"]
                    urlTest = requests.get(url)
                    if urlTest.status_code == 200:
                        return 1
        except Exception as e:
            print(metadata["title"] + "error")
            error_details = str(worksheet['H'+str(i)].value)
            worksheet['H'+str(i)] = error_details + ", " + str(e)
            print(e)
    
    return 0

for i in range(0,iteration):
    base_cell = 'A' + str(current_row)
    link = originalLink + str(worksheet[base_cell].value)
    metadata_file = requests.get(link)
    m_jsonResponse = metadata_file.json()
    
    metadata = m_jsonResponse["result"]
    
    #Update Frequency
    currentCell = 'C' + str(current_row)
    update_frequency_keywords = ["frequency","update","frequency-of-update","update-frequency"]
    frequency_flag = False
    for row in metadata["extras"]:
        if row["key"] in update_frequency_keywords:
            if row["value"] != "" and row["value"] != None:
                frequency_flag = True
                break
    
    if frequency_flag:
        worksheet[currentCell] = 0.4
    else:
        worksheet[currentCell] = 0
    
    #Formats
    currentCell = 'G' + str(current_row)
    formats = []
    for row in metadata["resources"]:
        if row["format"] not in formats:
            formats.append(row["format"])
    worksheet[currentCell] = str(formats)
    
    #API
    currentCell = 'E' + str(current_row)
    worksheet[currentCell] = verify_metadata_api(metadata=metadata,i=current_row)
    
    #CKAN
    currentCell = 'F' + str(current_row)
    worksheet[currentCell] = 0
    for row in metadata["extras"]:
        if row["key"] == "harvest_source_title":
            if "CKAN" in row["value"]:
                worksheet[currentCell] = 1
    
    #Machine-readable
    currentCell = 'D' + str(current_row)
    
    ## CKAN
    if (int(worksheet['F'+str(current_row)].value) == 1):
        if worksheet['E'+str(current_row)].value == 1:
            worksheet[currentCell] = 0.25
    elif ("WMS" in formats or "WFS" in formats):
        if int(worksheet['E'+str(current_row)].value) == 1:
            worksheet[currentCell] = 0.25
    else:
        worksheet[currentCell] = 0
    print(metadata["title"] + " dataset done.")
    
    current_row+=1
    spreadsheet.save("Addition.xlsx")
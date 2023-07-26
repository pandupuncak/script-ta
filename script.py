import json
import requests
import pygsheets
from openpyxl import load_workbook


datasetUrl = "https://katalog.data.go.id/dataset/"
originalLink = "https://katalog.data.go.id/api/action/package_show?id="
client = pygsheets.authorize(service_file='sdi-script-4b8593c0b4ed.json')

# spreadsheet = client.open('Spreadsheet Script Python')
spreadsheet = load_workbook(filename = 'Spreadsheet Script Python.xlsx')
# worksheet = spreadsheet.sheet1
worksheet = spreadsheet.get_sheet_by_name("Sheet1")
inp = input("Enter link:")
iterator = 556

reusability_3 = ["CSV","WMS","WFS","TXT", "JSON", "HTML", "GeoJSON"]
reusability_1 = ["PDF","XLS"]

def get_key_from_metadata_extras_SDI(metadata,key):
    for dict in metadata["extras"]:
        if dict["key"] == key:
            return dict["value"]
    
    return "None"

def check_reusability_format(formats):
    if formats == []:
        return 0
    
    for format in formats:
        if format in reusability_3:
            return 3
            
    return 1

def check_processability_format(formats):
    if formats == []:
        return 0
    
    for format in formats:
        if format in reusability_3:
            return 1
            
    return 0.2

def check_proprietary_format(formats):
    if formats == []:
        return 0
    
    for format in formats:
        if format in reusability_3:
            return 1
            
    return 0

file = ""
while inp != "None":
    try:
        if inp == "Save":
            print("Saving.")
            spreadsheet.save('Spreadsheet Script Python.xlsx')
            inp = input("Enter link: ")
        elif inp == "Skip":
            print("SKIPPING.")
            spreadsheet.save('Spreadsheet Script Python.xlsx')
            currentCell = "B" + str(iterator)
            worksheet[currentCell] = "NONE"
            
            inp = input ("Enter Title: ")    
            currentCell = "C" + str(iterator)
            worksheet[currentCell] = inp
                
            iterator += 1
            inp = input("Enter link: ")
    except Exception as e:
        print (e)
        inp = input("Please repeat: ")  
    else:
        try:
            link = originalLink+inp
            file = requests.get(link)
            jsonResponse = file.json()
            # dict_items = jsonResponse.items()
            # print(dict_items)
            metadata = jsonResponse["result"]
            currentCell = "B" + str(iterator)
            
            #URL
            currentDatasetURL = datasetUrl + metadata["name"]
            # worksheet.update_value(currentCell, currentDatasetURL)
            worksheet[currentCell] = currentDatasetURL
            
            #Nama Dataset
            currentCell = 'C' + str(iterator)
            datasetName = metadata["title"]
            # worksheet.update_value(currentCell,datasetName)
            worksheet[currentCell] = datasetName
            
            #Kategori
            currentCell = 'D' + str(iterator)
            # worksheet.update_value(currentCell, get_key_from_metadata_extras_SDI(metadata,"kategori"))
            worksheet[currentCell] = get_key_from_metadata_extras_SDI(metadata,"kategori")
            
            #D1
            currentCell = 'I' + str(iterator)
            if metadata["notes"] != None:
                if metadata["notes"] != "" and metadata["notes"].lower() != metadata["title"].lower():
                    # worksheet.update_value(currentCell,1)
                    worksheet[currentCell] = 1
                else:
                    # worksheet.update_value(currentCell,0)
                    worksheet[currentCell] = 0
            else:
                worksheet[currentCell] = 0  
                
            #D2
            currentCell = 'J' + str(iterator)
            if metadata["tags"] != []:
                # worksheet.update_value(currentCell,1)
                worksheet[currentCell] = 1
            else:
                # worksheet.update_value(currentCell,0)
                worksheet[currentCell] = 0
                
            #File Format
            formats = []
            for resource in metadata["resources"]:
                if resource["format"] not in formats:
                    formats.append(resource["format"])
            
            #D3
            currentCell = 'K' + str(iterator)
            if metadata['url'] != None:
                if "https" in metadata["url"]:
                    # worksheet.update_value(currentCell,1)
                    worksheet[currentCell] = 1
                else:
                    # worksheet.update_value(currentCell,0)
                    worksheet[currentCell] = 0
            else:
                # worksheet.update_value(currentCell,0)
                worksheet[currentCell] = 0
            
            #Reusability
            currentCell = 'L' + str(iterator)
            # worksheet.update_value(currentCell,check_reusability_format(formats))
            worksheet[currentCell] = check_reusability_format(formats)
            
            #WMS/WFS
            currentCell = 'O' + str(iterator)
            if "WMS" in formats or "WFS" in formats:
                # worksheet.update_value(currentCell,1)
                worksheet[currentCell] = 1
            else:
                # worksheet.update_value(currentCell,0)
                worksheet[currentCell] = 0
            
            #Last Update
            currentCell = 'W' + str(iterator)
            if metadata["metadata_modified"] != "":
                # worksheet.update_value(currentCell,1)
                worksheet[currentCell] = 1
            else:
                # worksheet.update_value(currentCell,0)
                worksheet[currentCell] = 0
            
            #Machine-processable
            currentCell = 'Y' + str(iterator)
            # worksheet.update_value(currentCell,check_processability_format(formats))
            worksheet[currentCell] = check_processability_format(formats)
            
            #Non-proprietary
            currentCell = 'AA' + str(iterator)
            # worksheet.update_value(currentCell,check_proprietary_format(formats))
            worksheet[currentCell] = check_proprietary_format(formats)
            
            #Licenses
            currentCell = 'AC' + str(iterator)
            # worksheet.update_value(currentCell,metadata["license_title"])
            worksheet[currentCell] = metadata["license_title"]
            
            currentCell = 'AB' + str(iterator)
            if metadata['license_id'] == 'cc-by':
                # worksheet.update_value(currentCell,1)
                worksheet[currentCell] = 1
            else:
                # worksheet.update_value(currentCell,0)
                worksheet[currentCell] = 0
            
            currentCell = 'AF' + str(iterator)
            # worksheet.update_value(currentCell,inp)
            worksheet[currentCell] = inp
            
            print(datasetName +" evaluated.")
                
                
        except Exception as e:
            spreadsheet.save('Spreadsheet Script Python.xlsx')
            currentCell = 'B' + str(iterator)
            # worksheet.update_value(currentCell,"ERROR")
            print(e)
        finally:
            iterator += 1
            inp = input("Enter link:")
            link = originalLink + inp

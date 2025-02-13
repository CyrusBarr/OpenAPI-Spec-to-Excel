import json
import openpyxl
import sys
import yaml

def convertYAML(ymlswaggerfile):
    with open(ymlswaggerfile, 'r', encoding='utf-8') as yaml_in:
        yaml_object = yaml.safe_load(yaml_in)
        jsonswaggerfile=json.dumps(yaml_object, default=str)
        jsonswaggerfile=json.loads(jsonswaggerfile)
    return jsonswaggerfile

def readFile(jsonswaggerfile):
    if isinstance(jsonswaggerfile,str):
        with open(jsonswaggerfile,'r',  encoding='utf-8') as f:
            jsonswaggerfile=json.load(f)
            
    endpoints = []
    for path,methods in jsonswaggerfile.get("paths", {}).items():
        for method, description in methods.items():
            try:
                endpoints.append({'Path': path,
                                  'Method': method.upper(),
                                  'Description': description.get('description','N/A')})
            except AttributeError as e:
                continue
    app_name=jsonswaggerfile.get("info", {}).get("title", {})
    app_name=app_name[:26]
    workbook=openpyxl.Workbook()
    sheet=workbook.active
    sheet.title=f'{app_name}.xlsx'

    headers = ['Method','Path','Description']
    sheet.append(headers)
    for endpoint in endpoints:
        sheet.append([endpoint['Method'],endpoint['Path'],endpoint['Description']])
    
    try:
        workbook.save(f'{app_name}.xlsx')
        print(f"Save to file successfully as {app_name}.xlsx")
    except Exception as e:
        print(f'Error saving file..\n{e}')


if __name__ == "__main__":
    if len(sys.argv)<2 or len(sys.argv)>2:
        print("Wrong Usage")

swaggerfile = sys.argv[1]
if swaggerfile.lower().endswith('.json'):
    readFile(swaggerfile)
elif swaggerfile.lower().endswith(('yml','yaml')):
    readFile(convertYAML(swaggerfile))
else:
    print("File is of incorrect format..")

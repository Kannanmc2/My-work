import pandas
import json
 
excel_data_df = pandas.read_excel('test.xlsx', header=[0,1,2])
thisisjson = excel_data_df.to_json(orient='records')
print(thisisjson)

with open('m.json', 'w', encoding = 'utf-8') as json_file_handler:
        
        json_file_handler.write(json.dumps(thisisjson))

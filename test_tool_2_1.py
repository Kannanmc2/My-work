import pandas
import pandas as pd
import json
from openpyxl import load_workbook

count=0


def Excel_Readers():
    Read_excel_data = pandas.read_excel('Sample.xlsx', sheet_name='Company Details')
    Read_excel_master= pandas.read_excel('Sample.xlsx',sheet_name='Entity Master')
    return Read_excel_data,Read_excel_master;

Read_excel_data,Read_excel_master = Excel_Readers()

def Null_checker():
    Null_identfy =Read_excel_data.isnull().any().any()
    Null_identfy_mas =Read_excel_master.isnull().any().any()
    #return Null_identfy,Null_identfy_mas;



    def Show_Null_place():
    
        if(Null_identfy==True):
            Null_OB=Read_excel_data[Read_excel_data.isna().any(axis=1)]
            print('\n Data Sheet have a null values:\n',Null_OB)

            change_value=input('\n Enter the correct value:\n')
            Read_excel_data.fillna(change_value, limit = 1,inplace = True)
            print(Read_excel_data)
            Null_checker()

        if(Null_identfy_mas==True):
            Null_OBB=Read_excel_master[Read_excel_master.isna().any(axis=1)]
            print('Master Sheet have a Null values: ',Null_OBB)
            
            change_value=input('\n Enter the correct value:')
            Read_excel_master.fillna(change_value, limit = 1,inplace = True)
            print(Read_excel_master)
            Null_checker()

    Show_Null_place()
    

Null_checker()

def Identify_duplicates():
    
        def remove_duplicates():
            
                duplicat = Read_excel_data[Read_excel_data.duplicated(subset='Company_Name')]
                if(len(duplicat) !=0):
                   print("\n Same company_Name is detected:\n",duplicat)
                   
                else:
                    duplicat = Read_excel_data[Read_excel_data.duplicated(subset='PAN')]
                    if(len(duplicat)!=0):
                       print("\n Same PAN Number is detected:\n",duplicat)
                    else:
                        duplicat = Read_excel_data[Read_excel_data.duplicated(subset='CIN')]
                        if(len(duplicat)!=0):
                           print("\n Same CIN Number is detected:\n",duplicat)
                        else:
                            duplicat = Read_excel_data[Read_excel_data.duplicated(subset='GST')]
                            if(len(duplicat)!=0):
                               print("\n Same GST Number is detected:\n",duplicat)
                            else:
                                duplicat = Read_excel_data[Read_excel_data.duplicated(subset='Email')]
                                if(len(duplicat)!=0):
                                   print("\n Same Email is detected:\n",duplicat)
                return duplicat;   
                
        
        duplicat = remove_duplicates()



        def Data_sheet_COM_Master_sheet():
            
            duplicat_mas = Read_excel_master[Read_excel_master.duplicated()]
            if(len(duplicat) == 0 and len(duplicat_mas) == 0):
                df1 = pd.DataFrame(Read_excel_data, columns = ['Entity'])
                df2 = pd.DataFrame(Read_excel_master, columns = ['Entity'])
                print('\n')
                listt=df2['Entity'].values.tolist()
                compare =(Read_excel_data[Read_excel_data.Entity.isin(listt)==False])
                #compare =(Read_excel_data[Read_excel_data.Bank_Name.isin(listt)==True])
                if (len(compare) == 0):
                    #Read_excel = pandas.read_excel('Sample.xlsx', sheet_name='Bank Details')
                    Convert_json = Read_excel_data.to_json(orient='records')
                    print('\n Excel to Json conver successfuly \n')
                    print(Convert_json)
                    print('\n')

                    with open('kannan.json', 'w', encoding = 'utf-8') as json_file_handler:
                     json_file_handler.write(json.dumps(Convert_json, indent = 5))
                else:
                      print('\n Your Bank Name is not correct :')
                      print(compare.Entity)
                      
                      
        Data_sheet_COM_Master_sheet()
       
    
Identify_duplicates()




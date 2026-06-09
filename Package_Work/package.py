import pandas as pd
#from db_connectionlive import get_connection
#from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from datetime import datetime
from database import result_df1



result_df1['PACKAGEID'] = result_df1['PREV_ID'] + 1

start_id = int(result_df1['PACKAGEID'].iloc[0])


file_path = r"C:\Users\software.support\Desktop\New folder\Sample-Package.xlsx"
#df = pd.read_excel(file_path, sheet_name="Bulk PKG Sheet")
# Read first worksheet
sheet1 = pd.read_excel(file_path, sheet_name='BulkPKGSheet')
package_code = sheet1['PKG Code']
PACKAGE_NAME = sheet1['Final PKG Name']
Days = sheet1['Days / LOS']
DayCare= sheet1['Day Care']
EcoCeling = sheet1['Eco Celing']
TwinCeiling = sheet1['Twin Ceiling']
SingleCeling = sheet1['Single Celing']
Deluxeceiling =sheet1['Deluxe ceiling']
# Read second worksheet
sheet2 = pd.read_excel(file_path, sheet_name='BulkPKGComponent')
# Add required columns with default values
sheet1['PACKAGEID'] = range(start_id, start_id + len(sheet1))
sheet1['PACKAGECODE'] = ('Surgical'+ package_code ) 
sheet1['PACKAGENAME'] = PACKAGE_NAME
sheet1['APPLICABLEVISITTYPE'] = 109
sheet1['AUTOCOMPONENTORDERPLACING'] = 0
sheet1['APPLICABLEGENDER'] = 530
sheet1['ISMULTIVISITPACKAGE'] = 0
sheet1['PACKAGEABBREVIATION'] = ''
sheet1['STARTDATE'] = ''
sheet1['CHARGEABLE'] = 0
sheet1['PACKAGECATEGORY'] = 1052127
sheet1['PACKAGETYPE'] = 1052130
sheet1['ENDDATE'] = ''
sheet1['NOOFALLOWABLEVISITS'] = ''
sheet1['REMARKS'] = ''
sheet1['ACTIVE'] = 1
sheet1['DURATION'] = Days
sheet1['CREATEDBY'] = 19413005
sheet1['CREATEDDATETIME'] = datetime.today().strftime('%d-%m-%y')
sheet1['UPDATEDBY'] = ''
sheet1['UPDATEDDATETIME'] = ''
sheet1['HIBVERSION'] = 0

# Dynamic ServiceMasterID
sheet1['SERVICEMASTERID'] = range(21794, 21794 + len(sheet1))

sheet1['MAINPROCEDURE'] = 0
sheet1['NOOFDAYSFORPACKAGE'] = 0
sheet1['SHOWONWEBSITE'] = 0

# Required column sequence
required_columns = [
    'PACKAGEID',
    'PACKAGECODE',
    'PACKAGENAME',
    'APPLICABLEVISITTYPE',
    'AUTOCOMPONENTORDERPLACING',
    'APPLICABLEGENDER',
    'ISMULTIVISITPACKAGE',
    'PACKAGEABBREVIATION',
    'STARTDATE',
    'CHARGEABLE',
    'PACKAGECATEGORY',
    'PACKAGETYPE',
    'ENDDATE',
    'NOOFALLOWABLEVISITS',
    'REMARKS',
    'ACTIVE',
    'DURATION',
    'CREATEDBY',
    'CREATEDDATETIME',
    'UPDATEDBY',
    'UPDATEDDATETIME',
    'HIBVERSION',
    'SERVICEMASTERID',
    'MAINPROCEDURE',
    'NOOFDAYSFORPACKAGE',
    'SHOWONWEBSITE'
]

# Create final dataframe
output_df = sheet1[required_columns]

# Output file
output_file = r"C:\Users\software.support\Desktop\New folder\Output-Package.xlsx"


with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    output_df.to_excel(writer, sheet_name="Package", index=False)  
# Export to Excel
#output_df.to_excel(output_file, index=False)

#print("Excel file generated successfully.")
#print("Output File :", output_file)

# print("\nSheet2 Data:")
# print(sheet2)
sheet1['SERVICE_CODE'] = ('Surgical'+ package_code ) 
sheet1['SERVICE_NAME'] = PACKAGE_NAME
sheet1['OP'] =EcoCeling
sheet1['Day Care'] =DayCare
sheet1['Eco Celing'] =EcoCeling
sheet1['Twin Ceiling']=TwinCeiling
sheet1['Single Celing']=SingleCeling
sheet1['Deluxe ceiling']=Deluxeceiling
required_columns1 = [
    'SERVICE_CODE',
    'SERVICE_NAME',
    'OP',
    'Day Care',
    'Eco Celing',
    'Twin Ceiling',
    'Single Celing',
    'Deluxe ceiling'

]
output_df1 = sheet1[required_columns1]

# Export to Excel
output_file = r"C:\Users\software.support\Desktop\New folder\Output-Package.xlsx"

with pd.ExcelWriter(output_file,engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:

    output_df1.to_excel(writer, sheet_name='packagerate', index=False)
     #service_df.to_excel(writer, sheet_name="Service", index=False)

#output_df.to_excel(output_file, index=False)
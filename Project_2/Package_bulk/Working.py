import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from datetime import datetime
from db_connection import get_connection
from package import file_path
from service import sheet1
from Requiredparameter import RService_code
import numpy as np

conn = get_connection()
if not conn:
    print("DB Connection Failed")
    exit()


# Convert to SQL IN format → '1001','1002','1003'

# Step 3: Create Query
query = f""" select groupid,groupname from groups where ACTIVE=1 """
result_df = pd.read_sql(query, conn)

conn.close()

#result_df['groupid'] = result_df['groupid'].astype(str).str.strip()

group_map = dict(
    zip(
        result_df['GROUPNAME'].astype(str).str.strip(),
        result_df['GROUPID']
    )
)

sheet2 = pd.read_excel(file_path, sheet_name='BulkPKGComponent')
package_code = sheet2['PKG code']
PACKAGE_NAME = sheet2['PKG Name']
Component_Name = sheet2['Component Name']
Remarks= sheet2['Remarks']
Type= sheet2['Type'] 
DayCare= sheet2['Day Care'] 
EcoCeling= sheet2['Eco Celing'] 
TwinCeiling= sheet2['Twin Ceiling'] 
SingleCeling= sheet2['Single Celing'] 
Deluxeceiling =sheet2['Deluxe ceiling']

# Read second worksheet


sheet2['PKG code'] ='Surgical' + sheet2['PKG code'].astype(str)
sheet2['PACKAGENAME'] = PACKAGE_NAME
sheet2['Component Name'] = Component_Name
sheet2['Groupid'] = sheet2['Component Name'].astype(str).str.strip().map(group_map)
sheet2['Group_Code'] = ( sheet2['Component Name']  .str.extract(r'\((\d+)\)')[0])
sheet2['Remarks'] = Remarks
sheet2['Type'] = Type
sheet2['Day Care'] = DayCare
sheet2['Eco Celing'] = EcoCeling
sheet2['Twin Ceiling'] = TwinCeiling
sheet2['Single Celing'] = SingleCeling
sheet2['Deluxe ceiling']=Deluxeceiling 
sheet2['Service Name'] = np.select(
    [
        sheet2['Remarks'] == 'OT Charges',
        sheet2['Remarks'] == 'Surgeon Charges',
        sheet2['Remarks'] == 'Anaesthesia Charges',
        sheet2['Remarks'] == 'Assistant Charges'
    ],
    [
        'OT-' + sheet2['PKG Name'].astype(str),
        'Surgery-' + sheet2['PKG Name'].astype(str),
        'Anaesthesia-' + sheet2['PKG Name'].astype(str),
        'Assi.Surg.-' + sheet2['PKG Name'].astype(str)
    ],
    default=''
)
# sheet2['Servicecode'] = (
#     sheet2['PKG code']
#     .astype(str)
#     .str[-4:]
# )

sheet2['Service_Code'] = (RService_code + sheet2['Group_Code'].astype(str)+ 'S'+  sheet2['PKG code'].astype(str).str[-5:])

#sheet2['Service_Code'] = ('INDACKOC' + sheet2['Group_Code'].astype(str)+ 'S'+  sheet2['PKG code'].astype(str).str[-5:])

required_remarks = [
   'OT Charges',
'Surgeon Charges',
'Anaesthesia Charges',
'Assistant Charges'

]

# Filter rows based on Remarks column
filtered_df = sheet2[
    sheet2['Remarks'].isin(required_remarks)
].copy()

# Required column order
required_columns = [
    'PKG code',
'PKG Name',
'Component Name',
'Groupid',
'Group_Code',
'Remarks',
'Type',
'Day Care',
'Eco Celing',
'Twin Ceiling',
'Single Celing',
'Deluxe ceiling',
'Service Name',
'Service_Code'
]

# Final output
#output_df = sheet2[required_columns]
output_df = filtered_df[required_columns]
# Export to Excel
output_file = r"C:\Users\software.support\Desktop\New folder\Output-Package.xlsx"

with pd.ExcelWriter(output_file,engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:

    output_df.to_excel(writer, sheet_name='Working', index=False)
     #service_df.to_excel(writer, sheet_name="Service", index=False)

#output_df.to_excel(output_file, index=False)

#print("Excel file generated successfully.")
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
#from database import result_df4
from db_connection import get_connection
import numpy as np
import warnings
warnings.filterwarnings("ignore")


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

# Read Excel file
file_path = r"C:\Users\software.support\Desktop\New folder\Output-Package.xlsx"

sheet1 = pd.read_excel(file_path, sheet_name='Package')

file_path1 = r"C:\Users\software.support\Desktop\New folder\Sample-Package.xlsx"

sheet2 = pd.read_excel(file_path1, sheet_name='BulkPKGComponent')

#package_id = sheet1['PACKAGEID']
Comp_name = sheet2['Component Name']
# type = sheet2['Type']
# daycare=sheet2['Day Care']
# eco=sheet2['Eco Celing']
# twin=sheet2['Twin Ceiling']
# Single=sheet2['Single Celing']
# Deluxeceiling=sheet2['Deluxe ceiling']

# Convert numeric columns safely
cols = ['Eco Celing', 'Twin Ceiling', 'Single Celing']

for col in cols:
    sheet2[col] = pd.to_numeric(sheet2[col], errors='coerce')

Package_id = dict(
    zip(sheet1['PACKAGECODE'],
        sheet1['PACKAGEID'])
)

# result_df4['Componentid'] = result_df4['PREV_ID'] + 1
# start_id = int(result_df4['Componentid'].iloc[0])

sheet2['ID'] =''
sheet2['COMPONENTCODE'] = ''
sheet2['COMPONENTTYPE'] = np.select(
    [ Comp_name == 'Implant(40)',
     Comp_name == 'DrugsandMaterials',
        ],
    [ '1052375','1052379'],default='1052373')
sheet2['COMPONENTNAME'] = Comp_name
sheet2['COMPONENTGROUPID'] = sheet2['Component Name'].astype(str).str.strip().map(group_map)
sheet2['COMPONENTID'] = '0'
sheet2['ISAUTOALLOCATION'] = '0'
sheet2['AMOUNTLIMIT'] = ''
sheet2['QTYLIMIT'] = np.where(
    (sheet2['Type'].astype(str).str.strip() == 'Qty') &
    (sheet2['Remarks'].fillna('').astype(str).str.strip() == '') &
    (sheet2['Eco Celing'].fillna(0) > 0) &
    (sheet2['Twin Ceiling'].fillna(0) > 0) &
    (sheet2['Single Celing'].fillna(0) > 0),

    sheet2['Eco Celing'].fillna(0).astype(int),

    ''
)
sheet2['ISEXCLUDED'] = np.select(
    [ sheet2['Component Name'] == 'Implant(40)'],
    [ '1'],
    default='0')
sheet2['LIST_INDEX'] = ''
sheet2['PACKAGEID'] = ('Surgical'+ sheet2['PKG code']).map(Package_id)
sheet2['EXEMPTION'] = '0'
sheet2['ACTIVE'] = '1'
sheet2['ISPACKAGECOMPONENT'] = '0'
sheet2['TARIFFCLASSTYPE'] = ''
sheet2['TARIFFCLASSVALUE'] = ''
sheet2['GENERIC'] = '1'
sheet2['COMPONENTPRICELIMIT'] = ''
sheet2['CPT_CODE'] = ''
sheet2['TARIFF_GROUP_ID'] = ''
sheet2['PERCENTAGE'] = '0'
sheet2['COMPONENTDISCOUNTEDPRICE'] = ''


Not_required_Comp_name = [
   'Investigations',
    'procedure charges(206)',
    'DrugsandMaterials'
]

# Exclude these rows
filtered_df = sheet2[
    ~sheet2['Component Name'].isin(Not_required_Comp_name)
].copy()




required_columns = [
    'ID',
'COMPONENTCODE',
'COMPONENTTYPE',
'COMPONENTNAME',
'COMPONENTGROUPID',
'COMPONENTID',
'ISAUTOALLOCATION',
'AMOUNTLIMIT',
'QTYLIMIT',
'ISEXCLUDED',
'LIST_INDEX',
'PACKAGEID',
'EXEMPTION',
'ACTIVE',
'ISPACKAGECOMPONENT',
'TARIFFCLASSTYPE',
'TARIFFCLASSVALUE',
'GENERIC',
'COMPONENTPRICELIMIT',
'CPT_CODE',
'TARIFF_GROUP_ID',
'PERCENTAGE',
'COMPONENTDISCOUNTEDPRICE'
]

output_df = filtered_df[required_columns]

# Export to Excel
output_file = r"C:\Users\software.support\Desktop\New folder\Output-Package.xlsx"

with pd.ExcelWriter(output_file,engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:

    output_df.to_excel(writer, sheet_name='Comp1', index=False)
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
#from database import result_df4
from db_connection import get_connection
from Tariff import output_file
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
    (sheet2['Remarks'].astype(str).str.strip() == '') &
    (sheet2['Eco Celing']> 0) &
    (sheet2['Twin Ceiling']> 0) &
    (sheet2['Single Celing']> 0),

    sheet2['Eco Celing'].astype(int),

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
#Comp2-----------------------------------------

sheet4 = pd.read_excel(file_path, sheet_name='Working')
sheet5 = pd.read_excel(file_path, sheet_name='Package')
sheet3 = pd.read_excel(file_path, sheet_name='Servicemaster')

#package_id = sheet1['PACKAGEID']
COMPONENTTYPE = sheet4['Component Name']
#Comp_name = sheet4['Service Name']
COMPONENTGROUPID = sheet4['Groupid']
#COMPONENT_ID = sheet3['SERVICE_MASTER_ID']


Package_id = dict(
    zip(sheet5['PACKAGECODE'],
        sheet5['PACKAGEID'])
)

COMPONENTID = dict(
    zip(sheet3['SERVICE_CODE'],
        sheet3['SERVICE_MASTER_ID'])
)


sheet4['ID'] =''
sheet4['COMPONENTCODE'] = ''
sheet4['COMPONENTTYPE'] = np.select(
    [ COMPONENTTYPE == 'IMPLANT 40',
     COMPONENTTYPE == 'DrugsAndMaterials',
        ],
    [ '1052375','1052379'],default='1052373')
sheet4['COMPONENTNAME'] = sheet4['Service Name']
sheet4['COMPONENTGROUPID'] = COMPONENTGROUPID
sheet4['COMPONENTID'] = sheet4['Service_Code'].astype(str).str.strip().map(COMPONENTID)
sheet4['ISAUTOALLOCATION'] = '0'
sheet4['AMOUNTLIMIT'] = ''
sheet4['QTYLIMIT'] = '1'
sheet4['ISEXCLUDED'] = '0'
sheet4['LIST_INDEX'] = ''
sheet4['PACKAGEID'] = sheet4['PKG code'].map(Package_id)
sheet4['EXEMPTION'] = '0'
sheet4['ACTIVE'] = '1'
sheet4['ISPACKAGECOMPONENT'] = '0'
sheet4['TARIFFCLASSTYPE'] = ''
sheet4['TARIFFCLASSVALUE'] = ''
sheet4['GENERIC'] = '1'
sheet4['COMPONENTPRICELIMIT'] = ''
sheet4['CPT_CODE'] = ''
sheet4['TARIFF_GROUP_ID'] = ''
sheet4['PERCENTAGE'] = '0'
sheet4['COMPONENTDISCOUNTEDPRICE'] = ''

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

output_df2 = sheet4[required_columns]

# Export to Excel
with pd.ExcelWriter(output_file,engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:

    output_df2.to_excel(writer, sheet_name='Comp2', index=False)

#Comp3-----------------------------------------

sheet2['Sr. No']=sheet2['Sr. No']
sheet2['PKG code']=('Surgical'+sheet2['PKG code'])
sheet2['PKG Name']=sheet2['PKG Name']
sheet2['Component Name']=sheet2['Component Name']
sheet2['COMPONENTGROUPID'] = sheet2['Component Name'].astype(str).str.strip().map(group_map)
sheet2 = sheet2[
    (sheet2['Type'] == 'Qty') &
    (sheet2['Remarks'].fillna('').str.strip() != '')
]
sheet2['Type']=sheet2['Type']
sheet2['Day Care']=sheet2['Day Care']
sheet2['Eco Celing']=sheet2['Eco Celing']
sheet2['Twin Ceiling']=sheet2['Twin Ceiling']
sheet2['Single Celing']=sheet2['Single Celing']
sheet2['Deluxe ceiling']=sheet2['Deluxe ceiling']



required_columns = [
'Sr. No',
'PKG code',
'PKG Name',
'Component Name',
'COMPONENTGROUPID',
'Remarks',
'Type',
'Day Care',
'Eco Celing',
'Twin Ceiling',
'Single Celing',
'Deluxe ceiling',

]

output_df1 = sheet2[required_columns]

# Export to Excel

with pd.ExcelWriter(output_file,engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:

    output_df1.to_excel(writer, sheet_name='Comp3', index=False)

# Packagerate ----------------------------------------------------------------------------
    
df= pd.read_excel(file_path1, sheet_name='BulkPKGComponent')

# Room type columns
room_columns = ['Day Care', 'Eco Celing', 'Twin Ceiling', 'Single Celing','Deluxe ceiling']

filter_values = [
    'DrugsandMaterials',
    'Investigations',
    'procedure charges(206)'
]

df_filtered = df[
    df['Component Name'].isin(filter_values)
]

# Convert columns into rows
result = df_filtered.melt(
    id_vars=['PKG code', 'PKG Name', 'Component Name'],
    value_vars=room_columns,
    var_name='Room_Type',
    value_name='Amount'
)

# Remove blank/null values
result = result.dropna(subset=['Amount']).round(0)
#result['PKG code'] = 'Surgical' + result['PKG code'].astype(str)

# Optional: Remove zero values
result = result[result['Amount'] != 0]

# Final column order
result = result[[
    'PKG code',
    'PKG Name',
    'Component Name',
    'Room_Type',
    'Amount'
]]

#print(result)
result['PKG code'] = 'Surgical' + result['PKG code'].astype(str)
# Save output

with pd.ExcelWriter(output_file,engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:

    result.to_excel(writer, sheet_name='Packagew_rate', index=False)
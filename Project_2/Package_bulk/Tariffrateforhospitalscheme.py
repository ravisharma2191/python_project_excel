import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from db_connection import get_connection
from datetime import datetime

import numpy as np
import warnings
warnings.filterwarnings("ignore")
hospital_id = input("Enter  HOSPITAL_ID: ")
hospital_ids = [x.strip() for x in hospital_id.split(",")]

# Convert list to SQL format
hospital_ids_str = ",".join(hospital_ids)
# hospital_id = input("Enter  Tariff HOSPITAL_ID: ")
# tariffgroup_id = input("Tariffgroup_ID: ")
conn = get_connection()
if not conn:
    print("DB Connection Failed")
    exit()

query1 = f""" select name,id from orgstructure where isactive=1 and id in ( {hospital_ids_str}) """
result_df1 = pd.read_sql(query1, conn)

query3 = f""" SELECT id,prev_id,( id - prev_id ) AS cnt
FROM(SELECT id,LAG(id, 1) OVER( ORDER BY id) AS prev_id 
FROM tariff) list WHERE id - prev_id > 20000 FETCH FIRST 1 ROW ONLY """
result_df3 = pd.read_sql(query3, conn)

conn.close()


# GET VALUES
# =========================

# for _, row in result_df1.iterrows():

#     hospital_name = row['NAME']
#     hospital_id1 = row['ID']
#tariffgroup_name = result_df2.iloc[0]['TARIFFGROUP_NAME']
result_df3['tariffid'] = result_df3['PREV_ID'] + 1
start_id = int(result_df3['tariffid'].iloc[0])

# Path--------------------

file_path = r"C:\Users\software.support\Desktop\New folder\Service.xlsx"


df = pd.read_excel(file_path, sheet_name='IT')

# Room type columns
room_columns = ['OP',	'Day Care',	'Economy',	'Twin Sharing',	'Single',	'Deluxe',	'Suite Room',	'ICU',]


# Convert columns into rows
result = df.melt(
    id_vars=['SERVICE_CODE', 'SERVICE_NAME', 'SERVICE_MASTER_ID'],
    value_vars=room_columns,
    var_name='Room_Type',
    value_name='Amount'
)

# Remove blank/null values
result = result.dropna(subset=['Amount'])

# Optional: Remove zero values
result = result[result['Amount'] != 0]

# Final column order
result = result[[
    'SERVICE_CODE',
    'SERVICE_MASTER_ID',
    'SERVICE_NAME',
    'Room_Type',
    'Amount'
]]

#print(result)

# Save output
output_file = r"C:\Users\software.support\Desktop\New folder\Srvicerate.xlsx"

with pd.ExcelWriter(output_file,engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:

    result.to_excel(writer, sheet_name='work_rate', index=False)

sheet1 = pd.read_excel(output_file, sheet_name='work_rate')

Service_id = sheet1['SERVICE_MASTER_ID']
SERVICE_NAME = sheet1['SERVICE_NAME']
Room_Type= sheet1['Room_Type'] 
Amount= sheet1['Amount'] 

room_name = np.select(
    [  Room_Type =='OP',
    Room_Type =='Day Care',
    Room_Type =='Economy',
    Room_Type =='Twin Sharing',
    Room_Type =='Single',
    Room_Type =='Deluxe',
    Room_Type =='Suite Room',
    Room_Type =='ICU',

    ],
    [
       'OP',
'Day Care',
'Economy',
'Twin Sharing',
'Single',
'Deluxe',
'Suite Room',
'ICU',

    ],
    default=''
)

#Sheet Column parameter-----------------

final_list = []

for _, row in result_df1.iterrows():

    hospital_name = row['NAME']
    hospital_id1 = row['ID']

    temp_sheet = sheet1.copy()

    # -------------------------------------

    temp_sheet['ID'] = range(start_id, start_id + len(temp_sheet))
    temp_sheet['ORIGINALID'] = range(start_id, start_id + len(temp_sheet))

    start_id += len(temp_sheet)

    # -------------------------------------

    temp_sheet['TARIFFVERSION'] = '0'

    temp_sheet['NAME'] = np.where(
        room_name != 'OP',

        temp_sheet['SERVICE_NAME'].astype(str)
        + "/"
        + hospital_name
        + "/"
        + room_name,

        temp_sheet['SERVICE_NAME'].astype(str)
        + "/"
        + hospital_name
    )

    temp_sheet['TARIFFGROUP'] = 29385
    temp_sheet['TOTALCHARGES'] = temp_sheet['Amount']
    temp_sheet['ADVANCE'] = 0
    temp_sheet['HOSPITALID'] = hospital_id1
    temp_sheet['CREATEDBY'] = 19413005
    temp_sheet['UPDATEDBY'] = ''
    temp_sheet['CREATEDDATETIME'] = datetime.today().strftime('%d-%m-%y')
    temp_sheet['UPDATEDDATETIME'] = ''
    temp_sheet['ACTIVE'] = 1
    temp_sheet['TARIFFTYPE'] = 729645
    temp_sheet['SERVICE_ID'] = temp_sheet['SERVICE_MASTER_ID']
    temp_sheet['TARIFFGROUPID'] = ''

    # -------------------------------------

    temp_sheet['TARFFIC_CLASS_TYPE'] = np.select(
        [
            temp_sheet['Room_Type'] == 'OP'
        ],
        [
            'HOSPITAL'
        ],
        default='TARIFF_GROUP'
    )

    # -------------------------------------

    temp_sheet['EFFECTIVEDATE'] = datetime.today().strftime('%d-%m-%y')
    temp_sheet['ISFIXED'] = 'N'
    temp_sheet['VALIDITY'] = '0'
    temp_sheet['ADVANCETYPE'] = '0'

    # -------------------------------------

    temp_sheet['TARFFIC_CLASS_ID'] = np.select(
        [
            temp_sheet['Room_Type'] == 'OP',
            temp_sheet['Room_Type'] == 'Day Care',
            temp_sheet['Room_Type'] == 'Economy',
            temp_sheet['Room_Type'] == 'Twin Sharing',
            temp_sheet['Room_Type'] == 'Single',
            temp_sheet['Room_Type'] == 'Deluxe',
            temp_sheet['Room_Type'] == 'Suite Room',
            temp_sheet['Room_Type'] == 'ICU',
        ],
        [
            str(hospital_id1),
            '1742381',
            '1742374',
            '1742375',
            '1742377',
            '1742378',
            '1742379',
            '1742380'
        ],
        default=''
    )

    # -------------------------------------

    temp_sheet['PERCENTAGE_TARIFF'] = ''
    temp_sheet['CHARGE_PERHOUR'] = '0'
    temp_sheet['DEPARTMENT_ID'] = ''
    temp_sheet['ALIASCODE'] = ''
    temp_sheet['ALIASNAME'] = ''
    temp_sheet['NEWTARIFFGROUPID'] = ''
    temp_sheet['SPONSOR_ID'] = ''

    # -------------------------------------

    final_list.append(temp_sheet)

# =========================================
# MERGE ALL HOSPITAL DATA
# =========================================

output_df = pd.concat(final_list, ignore_index=True)

# Required column order
required_columns = [
'ID',
'ORIGINALID',
'TARIFFVERSION',
'NAME',
'TARIFFGROUP',
'TOTALCHARGES',
'ADVANCE',
'HOSPITALID',
'CREATEDBY',
'UPDATEDBY',
'CREATEDDATETIME',
'UPDATEDDATETIME',
'ACTIVE',
'TARIFFTYPE',
'SERVICE_ID',
'TARFFIC_CLASS_TYPE',
'TARFFIC_CLASS_ID',
'EFFECTIVEDATE',
'ISFIXED',
'VALIDITY',
'ADVANCETYPE',
'TARIFFGROUPID',
'PERCENTAGE_TARIFF',
'CHARGE_PERHOUR',
'DEPARTMENT_ID',
'ALIASCODE',
'ALIASNAME',
'NEWTARIFFGROUPID',
'SPONSOR_ID'
]

output_df = output_df[required_columns]

# Export to Excel
output_file = r"C:\Users\software.support\Desktop\New folder\Srvicerate.xlsx"

with pd.ExcelWriter(output_file,engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:

    output_df.to_excel(writer, sheet_name='final_rate', index=False)



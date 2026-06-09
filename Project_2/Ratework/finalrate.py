import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from DbConfig.db_connection import get_connection
from Requiredparameter import hospital_id,tariffgroup_id,service_center_id,department_id
import numpy as np
from datetime import datetime
import warnings
warnings.filterwarnings("ignore")

conn = get_connection()
if not conn:
    print("DB Connection Failed")
    exit()

query1 = f""" select name,id from orgstructure where isactive=1 and id='{hospital_id}' """
result_df1 = pd.read_sql(query1, conn)

query2 = f""" select tariffgroup_name,id from tariffgroup where id='{tariffgroup_id}' """
result_df2 = pd.read_sql(query2, conn)

query3 = f""" SELECT id,prev_id,( id - prev_id ) AS cnt
FROM(SELECT id,LAG(id, 1) OVER( ORDER BY id) AS prev_id 
FROM tariff) list WHERE id - prev_id > 20000 FETCH FIRST 1 ROW ONLY """
result_df3 = pd.read_sql(query3, conn)

query4 = f""" SELECT SERVICE_LOC_ID,prev_id,( SERVICE_LOC_ID - prev_id ) AS cnt
FROM(SELECT SERVICE_LOC_ID,LAG(SERVICE_LOC_ID, 1) OVER( ORDER BY SERVICE_LOC_ID) AS prev_id 
FROM servicelocationmap) list WHERE SERVICE_LOC_ID - prev_id > 2000
FETCH FIRST 1 ROW ONLY """
result_df4 = pd.read_sql(query4, conn)

conn.close()


# GET VALUES
# =========================

hospital_name = result_df1.iloc[0]['NAME']
tariffgroup_name = result_df2.iloc[0]['TARIFFGROUP_NAME']
result_df3['tariffid'] = result_df3['PREV_ID'] + 1
start_id = int(result_df3['tariffid'].iloc[0])




file_path = r"C:\Users\software.support\Desktop\New folder\Output-Package_rate.xlsx"


sheet1 = pd.read_excel(file_path, sheet_name='work_rate')

Service_id = sheet1['SERVICE_MASTER_ID']
SERVICE_NAME = sheet1['SERVICE_NAME']
Room_Type= sheet1['Room_Type'] 
Amount= sheet1['Amount'] 

room_name = np.select(
    [   Room_Type == 'Day Care',
        Room_Type == 'Single Celing',
        Room_Type == 'Twin Ceiling',
        Room_Type == 'Eco Celing',
        Room_Type == 'Deluxe ceiling'
    ],
    [
        'Day Care',
        'Single',
        'Twin Sharing',
        'Economy',
        'Deluxe'
    ],
    default=''
)

#Sheet Column parameter-----------------

sheet1['ID']  =range(start_id, start_id + len(sheet1))
sheet1['ORIGINALID']  =range(start_id, start_id + len(sheet1))
sheet1['TARIFFVERSION'] = '0'
sheet1['NAME'] = np.where(
    room_name != '',
    SERVICE_NAME.astype(str)
    + "/"
    + hospital_name
    + "/"
    + room_name
    + "/"
    + tariffgroup_name,

    SERVICE_NAME.astype(str)
    + "/"
    + hospital_name
    + "/"
    + tariffgroup_name
)
sheet1['TARIFFGROUP'] = 29385
sheet1['TOTALCHARGES'] = Amount
sheet1['ADVANCE'] = 0
sheet1['HOSPITALID'] = hospital_id
sheet1['CREATEDBY'] = 19413005
sheet1['UPDATEDBY'] = ''
sheet1['CREATEDDATETIME'] = datetime.today().strftime('%d-%m-%y')
sheet1['UPDATEDDATETIME'] = ''
sheet1['ACTIVE'] = 1
sheet1['TARIFFTYPE'] = 729645
sheet1['SERVICE_ID'] = Service_id
sheet1['TARIFFGROUPID'] = np.select(
    [ Room_Type == 'Day Care',
     Room_Type == 'Eco Celing',
    Room_Type ==  'Twin Ceiling',
    Room_Type == 'Single Celing',
    Room_Type == 'Deluxe ceiling'
        ],
    [ '1742381','1742374','1742375','1742377','1742378'],default='')
sheet1['TARFFIC_CLASS_TYPE'] = np.select(
    [ Room_Type == 'OP',],
    [ 'HOSPITAL'],default='CHARGECLASS')
sheet1['EFFECTIVEDATE'] = datetime.today().strftime('%d-%m-%y')
sheet1['ISFIXED'] = 'N'
sheet1['VALIDITY'] = '0'
sheet1['ADVANCETYPE'] = '0'
sheet1['TARFFIC_CLASS_ID'] = np.select(
    [ Room_Type == 'OP'
        ],
    [ hospital_id],default=tariffgroup_id)

sheet1['PERCENTAGE_TARIFF'] = ''
sheet1['CHARGE_PERHOUR'] = '0'
sheet1['DEPARTMENT_ID'] = ''
sheet1['ALIASCODE'] = ''
sheet1['ALIASNAME'] = ''
sheet1['NEWTARIFFGROUPID'] = ''
sheet1['SPONSOR_ID'] = ''

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

output_df = sheet1[required_columns]

# Export to Excel
output_file = r"C:\Users\software.support\Desktop\New folder\Output-Package_rate.xlsx"

with pd.ExcelWriter(output_file,engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:

    output_df.to_excel(writer, sheet_name='final_rate', index=False)



    #Location Sheet generate-------------------------------------------

result_df4['service_loc_id'] = result_df4['PREV_ID'] + 1
start_id2 = int(result_df4['service_loc_id'].iloc[0])

# hospital_id = input("Enter Location HOSPITAL_ID: ")
# department_id = input("Enter DEPARTMENT_ID: ")
# service_center_id = input("Enter SERVICE_CENTER_ID: ")

sheet7 = pd.read_excel(file_path, sheet_name='work_rate')

service_map = dict(
    zip(sheet7['SERVICE_CODE'],
        sheet7['SERVICE_MASTER_ID'])
)

sheet6 = pd.read_excel(file_path, sheet_name='Servicemaster')

sheet6['SERVICE_LOC_ID'] =range(start_id2, start_id2 + len(sheet6))
sheet6['HOSPITAL_ID'] =hospital_id
sheet6['DEPARTMENT_ID'] =department_id
sheet6['SERVICE_CENTER_ID'] =service_center_id
sheet6['IS_PRIMARY_LOCATION'] ='1'
sheet6['OP_BILL_TYPE'] ='0'
sheet6['IP_BILL_TYPE'] ='0'
sheet6['SERVICE_ID'] =sheet6['SERVICE_CODE'].astype(str).str.strip().map(service_map)
sheet6['TARIFFID'] =''
sheet6['CREATEDBY'] =19413005
sheet6['CREATEDDT'] =datetime.today().strftime('%d-%m-%y')
sheet6['UPDATEDBY'] =''
sheet6['UPDATEDDT'] =''
sheet6['ISSUSPEND'] ='0'
sheet6['SUSPENDDATE'] =''
sheet6['ECONOMICCLASS'] =''

# Required column order
required_columns = [
'SERVICE_LOC_ID',
'HOSPITAL_ID',
'DEPARTMENT_ID',
'SERVICE_CENTER_ID',
'IS_PRIMARY_LOCATION',
'OP_BILL_TYPE',
'IP_BILL_TYPE',
'SERVICE_ID',
'TARIFFID',
'CREATEDBY',
'CREATEDDT',
'UPDATEDBY',
'UPDATEDDT',
'ISSUSPEND',
'SUSPENDDATE',
'ECONOMICCLASS'
]
output_df1 = sheet6[required_columns]

# Export to Excel

with pd.ExcelWriter(output_file,engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:

    output_df1.to_excel(writer, sheet_name='locationmaster', index=False)

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import sys
sys.path.append(r"E:\python")
from DbConfig.db_connectionlive import get_connection
import warnings
warnings.filterwarnings("ignore")


conn = get_connection()
if not conn:
    print("DB Connection Failed")
    exit()

query1 = f""" SELECT PACKAGEID,prev_id,( PACKAGEID - prev_id ) AS cnt
FROM(SELECT PACKAGEID,LAG(PACKAGEID, 1) OVER( ORDER BY PACKAGEID) AS prev_id 
FROM packagedefinition) list WHERE PACKAGEID - prev_id > 2000
FETCH FIRST 1 ROW ONLY """
result_df1 = pd.read_sql(query1, conn)

query2 = f""" SELECT SERVICE_MASTER_ID,prev_id,( SERVICE_MASTER_ID - prev_id ) AS cnt
FROM(SELECT SERVICE_MASTER_ID,LAG(SERVICE_MASTER_ID, 1) OVER( ORDER BY SERVICE_MASTER_ID) AS prev_id 
FROM servicemaster) list WHERE SERVICE_MASTER_ID - prev_id > 2000
FETCH FIRST 1 ROW ONLY """
result_df2 = pd.read_sql(query2, conn)

query3 = f""" SELECT SERVICE_LOC_ID,prev_id,( SERVICE_LOC_ID - prev_id ) AS cnt
FROM(SELECT SERVICE_LOC_ID,LAG(SERVICE_LOC_ID, 1) OVER( ORDER BY SERVICE_LOC_ID) AS prev_id 
FROM servicelocationmap) list WHERE SERVICE_LOC_ID - prev_id > 2000
FETCH FIRST 1 ROW ONLY """
result_df3 = pd.read_sql(query3, conn)

query4 = f""" SELECT id,prev_id,( id - prev_id ) AS cnt
FROM(SELECT id,LAG(id, 1) OVER( ORDER BY id) AS prev_id 
FROM packagecomponent) list WHERE id - prev_id > 20000
FETCH FIRST 1 ROW ONLY """
result_df4 = pd.read_sql(query4, conn)

query5 = f""" SELECT id,prev_id,( id - prev_id ) AS cnt
FROM(SELECT id,LAG(id, 1) OVER( ORDER BY id) AS prev_id 
FROM tariff) list WHERE id - prev_id > 20000 FETCH FIRST 1 ROW ONLY """
result_df5 = pd.read_sql(query5, conn)



conn.close()
result_df1['packageid'] = result_df1['PREV_ID'] + 1
result_df2['service_master_id'] = result_df2['PREV_ID'] + 1
result_df3['service_loc_id'] = result_df3['PREV_ID'] + 1
result_df4['component_id'] = result_df4['PREV_ID'] + 1
result_df5['tariffid'] = result_df5['PREV_ID'] + 1
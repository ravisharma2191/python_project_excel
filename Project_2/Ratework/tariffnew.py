import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from DbConfig.db_connection import get_connection

file_path = r"C:\Users\software.support\Desktop\New folder\Output-Package_rate.xlsx"

df = pd.read_excel(file_path, sheet_name='Servicemaster')

# Step 2: Prepare Employee Code list (clean)
codes = df['SERVICE_CODE'].dropna().astype(str).str.strip()

# Convert to SQL IN format → '1001','1002','1003'
sql_list = ",".join([f"'{c}'" for c in codes])

# Step 3: Create Query
conn = get_connection()
if not conn:
    print("DB Connection Failed")
    exit()
query = f""" select service_master_id,service_code from servicemaster where service_code  in ({sql_list}) """

result_df = pd.read_sql(query, conn)

conn.close()

# MAP SERVICE_MASTER_ID USING SERVICE_CODE
# -----------------------------------
# Clean DB columns
result_df.columns = result_df.columns.str.strip()

# -----------------------------------
# MERGE DATA
# -----------------------------------
df = df.merge(
    result_df,
    on='SERVICE_CODE',
    how='left'
)

# Debug
print(df.columns.tolist())

# -----------------------------------
# ROOM TYPE COLUMNS
# -----------------------------------
room_columns = [
    'OP',
    'Day Care',
    'Eco Celing',
    'Twin Ceiling',
    'Single Celing',
    'Deluxe ceiling'
]

# -----------------------------------
# MELT DATA
# -----------------------------------
result = df.melt(
    id_vars=[
        'SERVICE_CODE',
        'SERVICE_MASTER_ID',
        'SERVICE_NAME'
    ],
    value_vars=room_columns,
    var_name='Room_Type',
    value_name='Amount'
)

# -----------------------------------
# REMOVE NULL / ZERO
# -----------------------------------
result = result.dropna(subset=['Amount'])

result = result[result['Amount'] != 0]

# -----------------------------------
# FINAL COLUMN ORDER
# -----------------------------------
result = result[
    [
        'SERVICE_CODE',
        'SERVICE_MASTER_ID',
        'SERVICE_NAME',
        'Room_Type',
        'Amount'
    ]
]

#print(result)

# Save output
output_file = r"C:\Users\software.support\Desktop\New folder\Output-Package_rate.xlsx"

with pd.ExcelWriter(output_file,engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:

    result.to_excel(writer, sheet_name='work_rate', index=False)
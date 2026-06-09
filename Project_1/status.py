import pandas as pd
import sys
sys.path.append(r"E:\python")
from DbConfig.db_connectionlive import get_connection
#from hrsheet import output_path
import warnings
warnings.filterwarnings("ignore")
from copydata import source_file

# DB Connection
conn = get_connection()

if not conn:
    print("DB Connection Failed")
    exit()

cursor = conn.cursor()
# Read Excel
#file_path_mail=output_path
file_path = r"E:\python\Project_1\HR_SHEET\Output.xlsx"
df = pd.read_excel(file_path, sheet_name="Sheet3")
# Remove blank records
df = df[['Employee Code', 'Final_Approved_LWD']].dropna()

# Loop row by row
for index, row in df.iterrows():
    empno = str(row['Employee Code']).strip()
    # Convert date
    lwd = pd.to_datetime(str(row['Final_Approved_LWD']),dayfirst=True).strftime('%d-%m-%Y 18:00:00')

    try:

        # Update Employee Table
        employee_query = f"""
        UPDATE employee
        SET 
            EMP_STATUS = 76,
            UPDATEDDATETIME = TO_DATE('{lwd}','DD-MM-YYYY HH24:MI:SS'),
            INACTIVEDATETIME = TO_DATE('{lwd}','DD-MM-YYYY HH24:MI:SS'),
            UPDATEDBY = 30604558720,
            SHOWONWEBSITE = 0
        WHERE empno = '{empno}'
        """
        cursor.execute(employee_query)

        # Update employeecategorymap Table
        Categorymap_query = f"""
        delete from  employeecategorymap where groupid=760 and employeeid in (
        select employee_id from employee where empno ='{empno}')  """

        cursor.execute(Categorymap_query)

        deleted_count = cursor.rowcount
        #print(f"Employee Code: {empno} | Records Deleted from employeecategorymap: {deleted_count}")

        # Update HISUSER Table
        hisuser_query = f"""
        UPDATE hisuser SET isactive = 0 WHERE login = '{empno}' """

        cursor.execute(hisuser_query)

        print( f"Updated Successfully: {empno} | "f"Category Records Deleted: {deleted_count}")

    except Exception as e:
        print(f"Error for {empno} : {e}")

# Commit all updates
conn.commit()

print("All Records Updated Successfully")
# Close connection
cursor.close()
conn.close()
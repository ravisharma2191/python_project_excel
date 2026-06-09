import oracledb
import warnings
warnings.filterwarnings("ignore")

def get_connection():
    try:
        connection = oracledb.connect(
            user="shalbylive",
            password="ShalbyLive",
            dsn="10.18.0.61:1521/shalbypdb"
        )
       # print("Connected to Oracle DB")
        return connection

    except Exception as e:
        print("Connection Error:", e)
        return None
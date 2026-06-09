import oracledb

def get_connection():
    try:
        connection = oracledb.connect(
            user="shalby070426",
            password="shalby070426",
            dsn="10.0.0.21:1521/shalbypdb"
        )
        #print("Connected to Oracle DB")
        return connection

    except Exception as e:
        print("Connection Error:", e)
        return None
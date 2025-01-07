import pyodbc
from dotenv import load_dotenv
import os

# Memuat file .env
load_dotenv()

# Informasi koneksi
server = os.getenv("server")
database = os.getenv("database")
username = os.getenv("usernamesql")
password = os.getenv("password")
KodeDist = os.getenv("KodeDist")

# Membuat koneksi
try:
    connection = pyodbc.connect(
        f'DRIVER={{SQL Server}};SERVER={server};DATABASE={database};UID={username};PWD={password}'
    )
    print("Koneksi berhasil!")
    
    # Membuat cursor
    cursor = connection.cursor()
    
    # Query dengan parameter cust
    cursor.execute("SELECT * FROM fcustmst a WHERE a.KODECABANG = ?", (KodeDist,))

    # Query dengan parameter product
    cursor.execute("SELECT * FROM fmaster_dist WHERE DISTID = ?", (KodeDist,))

    # Query dengan parameter sales
    cursor.execute("SELECT * FROM fmap_salesman_dist")

    for row in cursor.fetchall():
        print(row)
    
    # Menutup koneksi
    connection.close()

except Exception as e:
    print("Terjadi kesalahan:", e)
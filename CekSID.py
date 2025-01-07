import pandas as pd
import pyodbc
from dotenv import load_dotenv
import os

# Memuat file .env
load_dotenv()

# Membaca variabel dari .env
InvoiceNo = os.getenv("InvoiceNo")
SalesNo = os.getenv("SalesNo")
CustNo = os.getenv("CustNo")
Pcode = os.getenv("Pcode")
KodeDist = os.getenv("KodeDist")
TypeInvoice = os.getenv("TypeINV")
FlagBonus = os.getenv('FlagBonus')

# Fungsi untuk memeriksa apakah nilai adalah 'null' atau kosong
def is_null(value):
    return value is None or value == '' or value.lower() == 'null'

def check_mapping_and_duplicates(file_path):
    # Membaca file Excel
    df = pd.read_excel(file_path)

    # Memastikan kolom yang diperlukan ada dalam file
    required_columns = [InvoiceNo, SalesNo, CustNo, Pcode, TypeInvoice]
    for col in required_columns:
        if col not in df.columns:
            print(f"Kolom yang dibutuhkan '{col}' tidak ditemukan dalam file!")
            return df
    
    # Menyiapkan list untuk hasil yang akan disimpan
    results = []

    # Membuat koneksi ke database
    try:
        connection = pyodbc.connect(
            f'DRIVER={{SQL Server}};SERVER={os.getenv("server")};DATABASE={os.getenv("database")};UID={os.getenv("usernamesql")};PWD={os.getenv("password")}'
        )

        # Membuat cursor
        cursor = connection.cursor()

        # Mendapatkan data dari tabel mappingan salesman
        cursor.execute("SELECT muid_dist FROM fmap_salesman_dist")
        salesman_data = cursor.fetchall()

        # Mendapatkan data dari tabel mappingan customer
        cursor.execute("SELECT CUSTNO_DIST FROM fcustmst_dist_map WHERE BRANCH_DIST = ?", (KodeDist,))
        customer_data = cursor.fetchall()

        # Mendapatkan data dari tabel mappingan product
        cursor.execute("SELECT PCODE FROM fmaster_dist WHERE DISTID = ?", (KodeDist,))
        product_data = cursor.fetchall()

        # Membuat kolom baru yang menggabungkan kolom yang diperlukan untuk memudahkan pengecekan
        df['combined'] = df[InvoiceNo].astype(str) + '-' + df[SalesNo].astype(str) + '-' + df[CustNo].astype(str) + '-' + df[Pcode].astype(str) + '-' + df[TypeInvoice].astype(str)

        # Mencari duplikat pada kolom 'combined' berdasarkan kondisi
        duplicates = df[df.duplicated(subset='combined', keep=False)]
        
        # Proses pengecekan SalesNo, CustNo, Pcode, dan duplikat
        for _, row in df.iterrows():
            # Pengecekan SalesNo
            if row[SalesNo] not in [x[0] for x in salesman_data]:
                results.append([row[InvoiceNo], row[SalesNo], row[CustNo], row[Pcode], KodeDist, "SalesNo tidak termapping"])

            # Pengecekan CustNo (memastikan perbandingan menggunakan format string)
            if str(row[CustNo]).strip() not in [str(x[0]).strip() for x in customer_data]:
                results.append([row[InvoiceNo], row[SalesNo], row[CustNo], row[Pcode], KodeDist, "CustNo tidak termapping"])

            # Pengecekan Pcode
            if row[Pcode] not in [x[0] for x in product_data]:
                results.append([row[InvoiceNo], row[SalesNo], row[CustNo], row[Pcode], KodeDist, "Pcode tidak termapping"])

            # Pengecekan duplikat dan FlagBonus
            if row['combined'] in duplicates['combined'].values:
                if row[TypeInvoice] == "INV":
                    flag_bonus_value = row.get(FlagBonus)
                    if is_null(flag_bonus_value):  # Jika FlagBonus tidak ada (null)
                        results.append([row[InvoiceNo], row[SalesNo], row[CustNo], row[Pcode], KodeDist, "Data double, FlagBonus tidak ada (null) untuk data"])
                    else:
                        results.append([row[InvoiceNo], row[SalesNo], row[CustNo], row[Pcode], KodeDist, "Data duplikat"])

        # Menutup koneksi
        connection.close()

        # Membuat DataFrame untuk hasil pengecekan
        results_df = pd.DataFrame(results, columns=['InvoiceNo', 'SalesNo', 'CustNo', 'Pcode', 'KodeDist', 'Status'])

        # Menyimpan hasil ke file Excel
        results_df.to_excel('hasil_cek_mapping_duplikat.xlsx', index=False)
        print("Hasil pengecekan telah disimpan dalam file 'hasil_cek_mapping_duplikat.xlsx'")

    except Exception as e:
        print("Terjadi kesalahan:", e)


# Masukkan path file .xlsx di bawah ini
file_path = 'NOV 2024 2024_MSLSINVDIST.xlsx'

# Langkah pertama: Pengecekan mapping (SalesNo, CustNo, Pcode) dan duplikat
check_mapping_and_duplicates(file_path)

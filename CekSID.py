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
    return value is None or value == '' or str(value).lower() == 'null'

def check_mapping_and_duplicates(file_path):
    # Membaca file Excel
    df = pd.read_excel(file_path)

    # Memastikan kolom yang diperlukan ada dalam file
    required_columns = [InvoiceNo, SalesNo, CustNo, Pcode, TypeInvoice]
    for col in required_columns:
        if col not in df.columns:
            print(f"Kolom yang dibutuhkan '{col}' tidak ditemukan dalam file!")
            return df
    
    results = []

    # Membuat koneksi ke database
    try:
        connection = pyodbc.connect(
            f'DRIVER={{SQL Server}};SERVER={os.getenv("server")};DATABASE={os.getenv("database")};UID={os.getenv("usernamesql")};PWD={os.getenv("password")}'
        )

        # Membuat cursor
        cursor = connection.cursor()

        # Mendapatkan data dari tabel mappingan
        cursor.execute("SELECT muid_dist FROM fmap_salesman_dist")
        salesman_data = [x[0] for x in cursor.fetchall()]

        cursor.execute("SELECT CUSTNO_DIST FROM fcustmst_dist_map WHERE BRANCH_DIST = ?", (KodeDist,))
        customer_data = [str(x[0]).strip() for x in cursor.fetchall()]

        cursor.execute("SELECT PCODE FROM fmaster_dist WHERE DISTID = ?", (KodeDist,))
        product_data = [x[0] for x in cursor.fetchall()]

        df['combined'] = df[InvoiceNo].astype(str) + '-' + df[SalesNo].astype(str) + '-' + df[CustNo].astype(str) + '-' + df[Pcode].astype(str) + '-' + df[TypeInvoice].astype(str)

        duplicates = df[df.duplicated(subset='combined', keep=False)]

        for _, row in df.iterrows():
            status_list = []

            # Pengecekan SalesNo
            if row[SalesNo] not in salesman_data:
                status_list.append("SalesNo tidak termapping")

            # Pengecekan CustNo
            if str(row[CustNo]).strip() not in customer_data:
                status_list.append("CustNo tidak termapping")

            # Pengecekan Pcode
            if row[Pcode] not in product_data:
                status_list.append("Pcode tidak termapping")

            # Pengecekan data double
            if row['combined'] in duplicates['combined'].values:
                if row[TypeInvoice] == "INV":
                    flag_bonus_value = row.get(FlagBonus)
                    if is_null(flag_bonus_value):
                        status_list.append("data double, FlagBonus tidak ada (null)")
                    else:
                        status_list.append("data double")

            # Gabungkan semua status dalam satu string
            if status_list:
                results.append([row[InvoiceNo], row[SalesNo], row[CustNo], row[Pcode], KodeDist, ", ".join(status_list)])

        connection.close()

        # Membuat DataFrame hasil
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

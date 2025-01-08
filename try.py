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

        # Mendapatkan data dari tabel mappingan sales
        cursor.execute("SELECT muid_dist, muid FROM fmap_salesman_dist")
        salesman_data = [(str(x[0]).strip(), x[1]) for x in cursor.fetchall()] 

        # Mendapatkan data dari tabel mappingan customer
        cursor.execute("SELECT CUSTNO_DIST, CUSTNO FROM fcustmst_dist_map WHERE BRANCH_DIST = ?", (KodeDist,))
        customer_data = [(str(x[0]).strip(), x[1]) for x in cursor.fetchall()]  # Tuple (CUSTNO_DIST, CUSTNO)

        # Mendapatkan data dari tabel mappingan product
        cursor.execute("SELECT PCODE, PCODE_PRC FROM fmaster_dist WHERE DISTID = ?", (KodeDist,))
        product_data = [(str(x[0]).strip(), x[1]) for x in cursor.fetchall()] 

        # Gabungkan kolom untuk pengecekan duplikat
        df['combined'] = df[InvoiceNo].astype(str) + '-' + df[SalesNo].astype(str) + '-' + df[CustNo].astype(str) + '-' + df[Pcode].astype(str) + '-' + df[TypeInvoice].astype(str)
        
        # Temukan duplikat
        duplicates = df[df.duplicated(subset='combined', keep=False)]

        for _, row in df.iterrows():
            status_list = []

            # Pengecekan SalesNo - cek apakah lebih dari satu data ditemukan
            matching_salesman = [(sls[1], sls[0]) for sls in salesman_data if sls[0].strip() == str(row[SalesNo]).strip()]
            if len(matching_salesman) == 0:
                status_list.append("SalesNo tidak termapping")
            elif len(matching_salesman) > 1:
                slsnos = [sls[0] for sls in matching_salesman]
                status_list.append(f"SalesNo termapping lebih dari 1, jumlah: {len(matching_salesman)} - SLSNO_PRC yang ditemukan: {', '.join(slsnos)}")

            # Pengecekan CustNo - cek apakah lebih dari satu data ditemukan
            matching_customer = [(cust[1], cust[0]) for cust in customer_data if cust[0].strip() == str(row[CustNo]).strip()]
            if len(matching_customer) == 0:
                status_list.append("CustNo tidak termapping")
            elif len(matching_customer) > 1:
                # Menampilkan semua CUSTNO yang cocok
                custnos = [cust[0] for cust in matching_customer]
                status_list.append(f"CustNo termapping lebih dari 1, jumlah: {len(matching_customer)} - CUSTNO_DIST yang ditemukan: {', '.join(custnos)}")

            # Pengecekan Pcode - cek apakah lebih dari satu data ditemukan
            matching_product = [(pcd[1], pcd[0]) for pcd in product_data if pcd[0].strip() == str(row[Pcode]).strip()]
            if len(matching_product) == 0:
                status_list.append("Pcode tidak termapping")
            elif len(matching_product) > 1:
                pcdnos = [pcd[0] for pcd in matching_product]
                status_list.append(f"Pcode termapping lebih dari 1, jumlah: {len(matching_product)} - PCODE_PRC yang ditemukan: {', '.join(pcdnos)}")

            # Pengecekan data double
            if row['combined'] in duplicates['combined'].values:
                if row[TypeInvoice] == "I":
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
file_path = 'DES_MSLSINVDIST.xlsx'

# Langkah pertama: Pengecekan mapping (SalesNo, CustNo, Pcode) dan duplikat
check_mapping_and_duplicates(file_path)

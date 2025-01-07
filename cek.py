import pandas as pd
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

def check_duplicates(file_path):
    # Membaca file Excel
    df = pd.read_excel(file_path)

    # Memastikan kolom yang diperlukan ada dalam file
    required_columns = [InvoiceNo, SalesNo, CustNo, Pcode, TypeInvoice]
    for col in required_columns:
        if col not in df.columns:
            print(f"Kolom yang dibutuhkan '{col}' tidak ditemukan dalam file!")
            return
    
    # Membuat kolom baru yang menggabungkan ketiga kolom tersebut untuk memudahkan pengecekan
    df['combined'] = df[InvoiceNo].astype(str) + '-' + df[SalesNo].astype(str) + '-' + df[CustNo].astype(str) + '-' + df[Pcode].astype(str) + '-' + df[TypeInvoice].astype(str)

    # Mencari duplikat pada kolom 'combined' berdasarkan kondisi
    duplicates = df[df.duplicated(subset='combined', keep=False)]

    if not duplicates.empty:
        print("Ditemukan data double:")
        for _, row in duplicates.iterrows():
            # Mengecek apakah TypeInvoice = 'INV' dan FlagBonus ada
            if row[TypeInvoice] == "INV":
                flag_bonus_value = row.get(FlagBonus)
                
                # Memeriksa jika FlagBonus tidak null atau kosong
                if not is_null(flag_bonus_value):  # Jika FlagBonus ada (tidak NaN)
                    if flag_bonus_value == "N":
                        print(f"Data duplikat: Invoice: {row[InvoiceNo]}, SalesID: {row[SalesNo]}, CustCode: {row[CustNo]}, ProductCode: {row[Pcode]}")
                    else:
                        # Hanya mencetak jika FlagBonus == 'N' (tidak duplikat)
                        pass
                else:
                    print(f"FlagBonus tidak ada (null) untuk data: {row[InvoiceNo]} - {row[SalesNo]} - {row[CustNo]} - {row[Pcode]}")
    else:
        print("Tidak ada data ganda.")

# Masukkan path file .xlsx di bawah ini
file_path = 'NOV 2024 2024_MSLSINVDIST.xlsx'
check_duplicates(file_path)

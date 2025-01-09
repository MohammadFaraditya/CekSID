import pandas as pd

# Daftar variabel header
InvoiceNo = 'invoiceNumber'
SalesNo = 'salesmanID'
CustNo = 'soldtoCustomerID'
Pcode = 'productCode'
KodeDist = 'DAMGT001'  # Tidak digunakan dalam data
TypeInvoice = 'sellingType'
FlagBonus = ''  # Kosong, diabaikan
Qty = 'qtySold'
dpp = 'lineGrossAmount'
tax = 'tax1'
nett = 'lineNetAmount'

# Fungsi untuk memanggil data hanya dengan header yang sesuai
def load_filtered_data(file_path):
    # Header yang diharapkan
    expected_headers = [InvoiceNo, SalesNo, CustNo, Pcode, TypeInvoice, Qty, dpp, tax, nett]
    
    try:
        # Membaca file Excel
        df = pd.read_excel(file_path)
        
        # Filter kolom yang hanya sesuai dengan header yang ada
        available_headers = [header for header in expected_headers if header in df.columns]
        
        # Menyimpan hanya kolom yang tersedia
        filtered_df = df[available_headers]
        
        print("Data berhasil dimuat dengan kolom berikut:")
        print(", ".join(available_headers))
        
        return filtered_df
    
    except FileNotFoundError:
        print(f"File '{file_path}' tidak ditemukan.")
        return None
    except Exception as e:
        print(f"Terjadi kesalahan: {e}")
        return None

# Contoh pemanggilan fungsi
data = load_filtered_data('Magetan.xlsx')

# Tampilkan data jika berhasil
if data is not None:
    print(data.head())

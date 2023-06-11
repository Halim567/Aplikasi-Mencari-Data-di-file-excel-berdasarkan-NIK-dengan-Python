import openpyxl
from dateutil.parser import parse

def cariDataExcel(file_path, nik):
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook.active

    for row_num, row in enumerate(sheet.iter_rows(min_row = 2, values_only = True), start=2):
        if row[0] == nik:
            try:
                parsed_tgl_lahir = parse(str(row[5]))
                format_tgl_lahir = parsed_tgl_lahir.strftime('%Y-%m-%d')
            except ValueError:
                format_tgl_lahir = ''
            print("=======================================================")
            print("NIK           :",f"{row[0]}")
            print("Nama          :",f"{row[1]}")
            print("Umur          :",f"{row[2]}","tahun")
            print("Tinggi Badan  :",f"{row[3]}","cm")
            print("Berat Badan   :",f"{row[4]}","kg")
            print("Tanggal Lahir :",f"{format_tgl_lahir}")
            print("Jenis Kelamin :",f"{row[6]}")
            print("=======================================================")
            print(f"Lokasi data berada dibaris ke {row_num} kolom ke A")
            break
    else :
        print("Data dengan NIK ",nik, " tidak ditemukan.")
        workbook.close()

print("--- PROGRAM MENCARI DATA BERDASARKAN NIK ---")
print(" ")
file_path = input("Masukan Nama file : ")
nik = int(input("Masukan NIK yang ingin dicari : "))
cariDataExcel(file_path, nik)
import tkinter as tk #Meng-import library UI
import os #Meng-import library untuk mengakses file
import openpyxl as xl #Meng-import library untuk membaca file excel
from dateutil.parser import parse #Meng-import library untuk format tanggal (optional)

#Inisiasi UI
windows = tk.Tk() #Inisiasi TKinter
windows.title("Aplikasi pencari data berdasarkan NIK") #Judul Program
windows.geometry("480x360") #Menentukan ukuran resolusi program
windows.resizable(False, False) #Mengunci resolusi program

#Deklarasi global variabel
sheet = None
text_error = None
text2 = None
text3 = None
text4 = None
text5 = None
text_loc = None
text_tgl = None

#Function untuk mengecek apakah file berada di dalam folder yang sama dengan program
def membuka_file():
    global sheet
    input_text = namaFile.get()
    if os.path.exists(input_text) and os.path.isfile(input_text): #Mengecek apakah file yang diinput berada didalam path/folder yang sama
        try:
            workbook = xl.load_workbook(input_text)
            sheet = workbook.active
            workbook.close()
            return True #Mengembalikan nilai true jika benar
        except:
            return False #Mengembalikan nilai false jika tidak

#Function untuk mengecek apakah file berhasil dibuka
def tombol():
    if membuka_file():
        tulisan = tk.Label(windows, text="File terbuka.\t\t\t", font=("Arial", 10))
        tulisan.place(x=75,y=60)
        tombol_klik1.config(state="normal") #Tombol kembali ke state normal
        nomorNik.config(state="normal")
    else:
        tulisan = tk.Label(windows, text="Gagal untuk membuka file!!", font=("Arial", 10), foreground='red')
        tulisan.place(x=75,y=60)

#Function untuk mengecek apakah nik yang diinputkan user ada di file excel
def pengecekan_nik():
    global sheet, text_error, text2, text3, text4, text5, text_loc, text_tgl
    nik = nomorNik.get() #Mengambil data yang diinput user dan menyimpannya di variabel nik
    tombol_reset.config(state="normal")
    tombol_klik1.config(state="disabled")
    #Melakukan perulangan untuk mendapatkan informasi tentang baris dan isi baris di file excel 
    for row_num, row in enumerate(sheet.iter_rows(min_row = 2, values_only = True), start=2):
        if str(row[0]) == str(nik): #Mengecek apakah nik yang diinput user ada di file
            try:
                parsed_tgl_lahir = parse(str(row[5])) #Mengambil info tentang format tanggal di kolom ke-5 excel yang kemudian disimpan di variabel persed_tgl_lahir
                format_tgl_lahir = parsed_tgl_lahir.strftime('%Y-%m-%d') #Mengubah format tanggal menjadi tahun - bulan - tanggal
            except ValueError:
                format_tgl_lahir = 'None' #Jika format tidak sesuai akan mengembalikan nilai 'None'

            text2 = tk.Label(windows, text="NIK           : " + str(row[0]))
            text2.place(x=25, y=160)
            text3 = tk.Label(windows, text="Tinggi badan : " + str(row[1]))
            text3.place(x=25, y=180)
            text4 = tk.Label(windows, text="Berat badan : " + str(row[2]))
            text4.place(x=25, y=200)
            text_tgl = tk.Label(windows, text="Tanggal Lahir : " + str(format_tgl_lahir))
            text_tgl.place(x=25, y=220)
            text5 = tk.Label(windows, text="Jenis kelamin : " + str(row[3]))
            text5.place(x=25, y=240)
            text_loc = tk.Label(windows, text="Lokasi data berada dibaris ke " + str(row_num))
            text_loc.place(x=25, y=260)
            try:
                text_error.destroy()
            except:
                text_error = None
            break
    else :
        text_error = tk.Label(windows, text="Data dengan NIK " + str(nik) +" tidak ditemukan!!", foreground='red')
        text_error.place(x=25, y=338)

#Function untuk mereset data
def reset():
    global text_tgl, text5, text2, text3, text4, text_error, text_loc
    tombol_reset.config(state="disabled")
    tombol_klik1.config(state="normal") 
    nomorNik.delete(0, tk.END) #Menghapus data yang ada di kotak area input user
    if text2:
        text2.destroy() #Menghapus data/value text
    if text3:
        text3.destroy()
    if text4:
        text4.destroy()
    if text5:
        text5.destroy()
    if text_error:
        text_error.destroy()
    if text_loc:
        text_loc.destroy()
    if text_tgl:
        text_tgl.destroy()

#Global varioabel untuk menampilkan text, tombol, dan garis
garis = tk.Canvas(windows, width=455) #Membuat garis
garis.place(x=14, y=-88)
garis.create_line(0, 180, 447, 180)
garis.create_line(0, 230, 180, 230)
garis.create_line(250, 230, 447, 230)

label_pack = tk.Label(windows, text="Aplikasi by Halim648", font=("Arial", 9)) #Membuat text untuk ditampilkan di apk
label_pack.place(x=355, y=338) #Meletakan text, tombol, dan kotak input sesuai dengan koordinat x dan y

judul1 = tk.Label(windows, text="Masukan nama file : ", font=("Arial", 11))
judul1.place(x=25, y=30)

judul2 = tk.Label(windows, text="Masukan no NIK     : ", font=("Arial", 11))
judul2.place(x=25, y=105)

judul3 = tk.Label(windows, text="Info Data")
judul3.place(x=205, y=130)

info1 = tk.Label(windows, text="Status : ", font=("Arial", 10))
info1.place(x=25,y= 60)

namaFile = tk.Entry(windows, width=25) #Membuat sebuah kotak area untuk user menginput data (data dalam bentuk string)
namaFile.place(x=164, y=33)

nomorNik = tk.Entry(windows, width=25, state="disabled")
nomorNik.place(x=164, y=108)

tombol_klik = tk.Button(windows, text="Kirim", command=tombol) #Membuat tombol untuk mengirim data yang diinput user ke program
tombol_klik.place(x=330, y=28)

tombol_klik1 = tk.Button(windows, text="Kirim", command=pengecekan_nik, state="disabled") #Tombol dalam status tidak bisa ditekan
tombol_klik1.place(x=330, y=103)

tombol_reset = tk.Button(windows, text="Reset", command=reset, state="disabled")
tombol_reset.place(x=380, y=103)

windows.mainloop() #Loop program
# Sertif Generator

Script python yang semoga berfaedah, terutama untuk sie dekdok dalam membuat sertifikat dan mendistribusikan e-sertif.

## Prasyarat
Install modul berikut terlebih dahulu:
1. `pandas`: Untuk membaca file spreadsheet\
instalasi: `pip install pandas`
2. `natsort`: Untuk mengurutkan item list dengan *natural sort*\
instalasi: `pip install natsort`

## Petunjuk penggunaan
### Dapatkan script
Dari *command line,* masuk ke direktori tempat script ini akan disimpan dan jalankan perintah\
`git clone https://gitlab.com/deutan/sertif-generator`\
Atau *download source code* dan ektrak pada direktori yang diinginkan.
### Yang perlu disiapkan sebelum eksekusi script
- File data penerima sertifikat dalam bentuk *spreadsheet* (.xls/.xlsx)
<<<<<<< HEAD
- Isi berkas memiliki kolom **Nama** dan **Email**. Contoh:
=======
- Isi berkas memiliki kolom **Nama** dan **Email**. Contoh:\
    
>>>>>>> 5e5b8c39e44f71a033634054927a3c9855998c1f
    | Nomor | *Nama* | *Email* | No. HP |
    | --- | --- | --- | --- |
    | 1 | Andi Budiarto | andi@mail.com | 0851234567890 |
    | 2 | Cici Cuita | cicici@cimail.com | 0812345678912 |
    | 3 | Dodi Edogawa | dodi@yuhuu.com | 0801234567892 |
- Template berupa gambar (.jpg/.png)
- [Opsional] File font (.ttf), lebih baik tipe *Monospace,* terutama jika letak nama dalam sertif diposisikan di tengah secara horizontal *(align: center)*.
- Izinkan akun gmail untuk [**Akses aplikasi yang kurang aman**](https://myaccount.google.com/lesssecureapps)
  
Semua file di atas sebaiknya diletakkan pada direktori yang sama dengan file script ini.

### Yang perlu diubah sebelum eksekusi script
Ubah value dari variabel-variabel berikut sesuai keperluan:
- `template = 'template.png'`\
  Isi dengan nama file gambar template.
- `spreadsheet_file = 'datalist.xlsx'`\
  Isi dengan nama file excel/spreadsheet berisi data penerima sertif.
- `name_font = ImageFont.truetype('RobotoMono.ttf', 150)`\
  Isi dengan Nama font atau *fullpath* alamat file font .ttf; argumen kedua merupakan ukuran font.
- `font_color = (71, 48, 149)`\
  Warna font dalam nilai RGB.
- `output_folder_name = 'Sertif Files'`\
  Nama folder yang akan dibuat secara otomatis sebagai tempat menyimpan file PDF sertif.
- `server = smtplib.SMTP(host='smtp.gmail.com', port=587)`\
  Host dan port dari penyedia layanan email. Script ini menggunakan gmail, ubah jika pengirim menggunakan layanan lain.
- Pada fungsi `insertNames()`, terdapat  kode berikut:\
  `x = (img_w - name_w)/2  # Center horizontally
  y = 970`\
  x dan y merupakan nilai dari posisi teks, diukur dari pojok kiri atas gambar template (jika nilai `x = 0` dan `y = 0`, teks akan berada di pojok kiri atas). Dalam script ini x akan menghitung posisi agar teks berada di tengah secara horizontal.
- Ubah struktur email: Nama pengirim, Judul, dan Isi pesan (dalam HTML).
    ``` python
    sender_name = 'Sender Name'
    subject = 'Mail Subject'

    body = ("""
    <html>
        <head></head>
        <body>
            <p>Hello!<br />
            How are you?<br />
            Here is the <a href="http://www.gitlab.com/deutandev/">link</a> you wanted.
            </p>
        </body>
    </html>
    """)
    ```

### Run!
Setelah siap semua, melalui *command line* masuk ke direktori tempat script disimpan. Lalu jalankan script dengan perintah\
`python sertif_generator.py`

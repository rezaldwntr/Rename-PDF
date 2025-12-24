========================================================================
PANDUAN PENGGUNAAN APLIKASI RENAME PDF MASSAL
========================================================================

DESKRIPSI:
Aplikasi ini berfungsi untuk mengubah nama (rename) file PDF dalam jumlah 
banyak secara otomatis berdasarkan daftar nama yang ada di file Excel.
Aplikasi ini AMAN karena tidak menimpa file asli, melainkan menyalin 
file yang sudah di-rename ke folder baru.

------------------------------------------------------------------------
A. CARA MEMBUAT APLIKASI (Hanya dilakukan sekali di awal)
------------------------------------------------------------------------
Jika Anda belum memiliki file ".exe" dan masih berupa script mentah, 
lakukan langkah ini:

1. Pastikan komputer terhubung ke Internet (untuk download library awal).
2. Pastikan file "PDF_Renamer_GUI.py" dan "Buat_Exe.bat" ada di 
   dalam satu folder yang sama.
3. Klik dua kali file "Buat_Exe.bat".
4. Tunggu proses berjalan (layar hitam) hingga muncul tulisan "SELESAI".
5. Akan muncul file baru bernama "Aplikasi_Rename_PDF.exe".
   File inilah yang bisa Anda pakai seterusnya atau dipindahkan ke 
   komputer lain.

------------------------------------------------------------------------
B. PERSIAPAN DATA (PENTING!)
------------------------------------------------------------------------
Sebelum menjalankan aplikasi, siapkan data Anda:

1. FILE EXCEL:
   - Buat file Excel baru.
   - Masukkan daftar nama baru di KOLOM A (Kolom pertama).
   - Pastikan urutan nama di Excel sesuai dengan urutan file PDF 
     (Aplikasi mengurutkan PDF berdasarkan Abjad/Angka: 01, 02, dst).
   - Tidak perlu menambahkan ".pdf" di belakang nama (aplikasi akan 
     menambahkannya otomatis).
   - Simpan file Excel dan tutup.

2. FILE PDF:
   - Pastikan semua file PDF terkumpul dalam satu folder.
   - Pastikan nama file PDF asli sudah bisa diurutkan (Sort by Name) 
     agar sesuai dengan urutan baris di Excel.

------------------------------------------------------------------------
C. CARA MENJALANKAN APLIKASI
------------------------------------------------------------------------
1. Klik dua kali "Aplikasi_Rename_PDF.exe".
2. Klik tombol "MULAI PROSES".
3. Jendela pertama: Pilih FOLDER tempat file PDF asli berada.
4. Jendela kedua: Pilih FILE EXCEL yang berisi daftar nama.
5. Tunggu proses berjalan.
6. Jika selesai, aplikasi akan memberitahu jumlah file yang berhasil.

------------------------------------------------------------------------
D. LOKASI HASIL
------------------------------------------------------------------------
Hasil file yang sudah di-rename akan berada di folder baru yang dibuat 
secara otomatis di samping folder asli Anda.
Nama foldernya diawali dengan: "HASIL_RENAME_..."

------------------------------------------------------------------------
PEMECAHAN MASALAH (TROUBLESHOOTING)
------------------------------------------------------------------------
1. "Windows protected your PC" saat membuka aplikasi?
   - Klik "More info", lalu klik tombol "Run anyway". 
   - Ini normal karena aplikasi ini dibuat sendiri dan tidak memiliki 
     sertifikat digital berbayar dari Microsoft.

2. Jumlah file tidak sesuai?
   - Cek apakah ada file PDF yang rusak (corrupt).
   - Cek apakah jumlah baris di Excel sama dengan jumlah file di folder.

3. Urutan nama tertukar?
   - Pastikan file PDF asli Anda memiliki penomoran yang konsisten 
     (Contoh: gunakan 01, 02, ... , 10. Jangan 1, 2, ... , 10 karena 
     komputer membaca 10 lebih dulu daripada 2).

========================================================================
Dibuat dengan Python
========================================================================

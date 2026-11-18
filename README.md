# Injeksi Kode HTML/JS melalui Atribut `embeddedHtml` pada Microsoft Word (.docx)

Repositori ini menyediakan skrip otomatisasi berbasis PowerShell untuk menguji dan mendemonstrasikan mekanisme injeksi kode HTML serta JavaScript ke dalam file dokumen Microsoft Word (`.docx`) melalui manipulasi atribut XML `embeddedHtml`, lengkap dengan panduan analisis keamanan dan mitigasinya.

> **⚠️ Catatan Edukasi & Audit:**  
> Dokumentasi dan skrip ini disusun untuk tujuan pembelajaran, pengujian keamanan (red teaming), dan penguatan pertahanan sistem (blue teaming). Gunakan hanya pada lingkungan lab atau sistem yang memiliki izin resmi.

---

## 📖 Latar Belakang & Mekanisme Celah

Tim peneliti keamanan siber dari Cymulate mengidentifikasi adanya perilaku unik pada fitur pemutar video daring (*Online Video*) di Microsoft Word (versi 2016 ke atas). Saat fitur ini digunakan, Microsoft Word menyematkan kode `<iframe>` pemutar video (seperti YouTube) ke dalam struktur internal dokumen.

Namun, alih-alih membatasi konten tersebut hanya untuk video resmi, mesin perender web di dalam Microsoft Word ternyata akan mengeksekusi kode HTML apa pun yang tertulis di dalam atribut `embeddedHtml`. Celah ini memungkinkan siapa saja untuk memodifikasi struktur XML dokumen `.docx` yang sudah memiliki video daring, menimpa nilai atribut `embeddedHtml` dengan skrip kustom, dan menjalankannya saat dokumen dibuka tanpa perlu menggunakan makro VBA.

Karena proses membongkar file `.docx`, menyunting XML secara manual, dan mengemasnya kembali cukup memakan waktu, skrip PowerShell di repositori ini hadir untuk mengotomatisasi seluruh tahapan tersebut.

---

## 🛠️ Panduan Penggunaan Skrip

### Prasyarat Sistem
Sebelum menjalankan skrip, pastikan lingkungan Anda memenuhi kebutuhan berikut:
1. **Dokumen Target:** Anda harus menyiapkan sebuah file Microsoft Word (`.docx`) yang sudah memiliki minimal satu video daring di dalamnya. Untuk membuatnya, buka Word lalu masuk ke menu **Insert > Online Video**, cari video apa saja, dan sisipkan ke dalam halaman.
2. **Versi PowerShell:** Skrip ini dibangun menggunakan **PowerShell 5.1**, sehingga Anda memerlukan PowerShell versi 5.1 atau yang lebih baru di sistem operasi Windows.

### Cara Memuat & Menjalankan Fungsi

Untuk memanggil fungsi `Inject-Docx` langsung melalui terminal PowerShell, lakukan teknik *dot-sourcing* terlebih dahulu agar fungsi dimuat ke dalam sesi terminal Anda:

```powershell
# Memuat skrip ke dalam sesi terminal
. C:\scripts\docx-embeddedhtml-injection.ps1

# Menjalankan fungsi injeksi
Inject-Docx -Path "C:\Lokasi\Dokumen\test.docx" -HtmlBlock "<h3>Uji Injeksi HTML</h3>"
```

Jika Anda ingin memanggil fungsi ini dari dalam skrip PowerShell lain, cukup sertakan baris berikut di bagian atas skrip Anda:

```powershell
. C:\scripts\docx-embeddedhtml-injection.ps1
```

### Rincian Parameter Fungsi

Sintaks lengkap untuk menjalankan fungsi ini adalah sebagai berikut:

```powershell
Inject-Docx -Path "C:\Lokasi\Dokumen\test.docx" -HtmlBlock "<h3>Uji Injeksi HTML</h3>" -DestinationName "dokumen_hasil.docx"
```

* **`-Path`** *(Wajib)*: Jalur lengkap (*full path*) menuju file dokumen `.docx` asli yang sudah berisi video daring.
* **`-HtmlBlock`** *(Wajib)*: Baris kode HTML atau JavaScript yang ingin disuntikkan ke dalam atribut `embeddedHtml`.
* **`-DestinationName`** *(Opsional)*: Nama file output untuk dokumen yang telah dimodifikasi. Jika parameter ini tidak diisi, sistem akan otomatis menyimpannya dengan nama bawaan `output.docx`.

---

## 🔍 Analisis Karakteristik Serangan

* **Vektor Serangan Remote (Jarak Jauh):** Penyerang tidak memerlukan akses fisik ke komputer target. File `.docx` yang telah dimodifikasi dapat dikirimkan dari jarak jauh melalui lampiran email (*phishing*), aplikasi pesan instan, atau tautan unduhan web.
* **Membutuhkan Interaksi Pengguna (*User Interaction*):** Kode yang disuntikkan tidak berjalan secara otomatis saat file diterima. Pengguna harus mengunduh dokumen, membukanya di Microsoft Word, menonaktifkan *Protected View* (jika aktif), dan mengeklik gambar pratinjau video agar kode HTML/JS tereksekusi secara lokal di bawah hak akses pengguna tersebut.
* **Sikap Resmi Vendor:** Hingga saat ini, Microsoft tidak mengklasifikasikan perilaku ini sebagai kerentanan keamanan (*vulnerability*) yang memerlukan pembaruan tambalan (*patch*). Dalam pernyataan resminya, Direktur Senior Microsoft (Jeff Jones) menyatakan bahwa produk tersebut menerjemahkan kode HTML sebagaimana mestinya sesuai rancangan awal (*working as designed*), sama seperti cara kerja perangkat lunak sejenis lainnya.

---

## 🛡️ Strategi Deteksi & Mitigasi (Panduan Defensif)

Karena tidak ada perbaikan resmi dari pihak vendor, tim keamanan dan administrator sistem perlu menerapkan kontrol berlapis untuk mendeteksi serta mencegah penyalahgunaan teknik ini:

### 1. Pemindaian Statis pada Email Gateway & DLP
Karena file `.docx` merupakan arsip terkompresi (ZIP), perangkat keamanan email dapat membongkar isinya dan memeriksa file `word/document.xml`.
* **Validasi Sumber Video:** Dokumen yang sah hanya memuat tag `<iframe>` yang mengarah ke domain resmi (seperti `youtube.com` atau `vimeo.com`). Blokir dokumen jika atribut `embeddedHtml` mengandung tag `<script>`, `<form>`, `javascript:`, atau pengkodean `data:text/html`.
* **Implementasi Aturan YARA:** Gunakan aturan YARA berikut pada sistem pemindai file untuk mendeteksi indikasi injeksi secara otomatis:

```yara
rule Detect_Docx_EmbeddedHtml_Injection {
    meta:
        description = "Mendeteksi potensi injeksi HTML/JS pada atribut embeddedHtml di file DOCX"
        author = "Security Audit Team"
        severity = "High"
    strings:
        \(zip_header = { 50 4B 03 04 }\)attr = "embeddedHtml=" ascii wide
        \$suspicious_1 = "<script" nocase ascii wide
        \$suspicious_2 = "<form" nocase ascii wide
        \$suspicious_3 = "javascript:" nocase ascii wide
        \$suspicious_4 = "data:text/html" nocase ascii wide
    condition:
        \$zip_header at 0 and attr and (any of (suspicious_*))
}
```

### 2. Penguatan Konfigurasi Endpoint (Hardening)
* **Blokir Video Daring via Group Policy (GPO):** Nonaktifkan fitur penyematan dan pemuatan konten video daring di Microsoft Office melalui pengaturan *Group Policy Object* untuk menutup celah ini secara menyeluruh di lingkungan organisasi.
* **Wajibkan Protected View:** Pastikan fitur *Protected View* selalu aktif untuk semua file yang diunduh dari internet atau berasal dari lampiran email eksternal agar pemuatan elemen web tertahan sebelum pengguna mengeklik *Enable Editing*.
* **Aktifkan Aturan Attack Surface Reduction (ASR):** Pada lingkungan Microsoft Defender for Endpoint, aktifkan aturan *"Block all Office applications from creating child processes"* guna mencegah `WINWORD.EXE` meluncurkan proses sistem berbahaya (seperti `powershell.exe`, `cmd.exe`, atau `mshta.exe`) apabila terjadi eksploitasi lanjutan.

---

## 📚 Sumber & Referensi
* Temuan asli oleh **Cymulate Research** terkait *Microsoft Word Online Video Embedded HTML Injection*.
* Repositori kode sumber asli: [thom-s/docx-embeddedhtml-injection](https://github.com).

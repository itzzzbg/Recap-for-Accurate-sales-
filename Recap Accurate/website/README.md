# Rekap Faktur Penjual

Website untuk menghitung jumlah No. Faktur per Nama Penjual dari PDF
"Laporan Penjualan All (Penjual)" ACCURATE. Faktur yang sama di beberapa
baris barang dihitung 1 kali. Hasil bisa diunduh sebagai Excel atau disalin.

PDF diproses di browser pengguna. Server hanya mengirim halaman, tidak
menerima atau menyimpan file laporan.

## Isi folder

| File | Fungsi |
|---|---|
| `index.html` | Halaman aplikasi |
| `vendor/` | pdf.js 3.11.174 dan ExcelJS 4.4.0 (lokal, tidak butuh internet) |
| `server.py` | Web server kecil, hanya pakai Python bawaan |
| `Jalankan Server.bat` | Double-click untuk menjalankan `server.py` di Windows |

## Menjalankan

**Pakai server.py (Windows/Linux, perlu Python 3):**

```
python server.py          # port 8080
python server.py 9000     # port lain
```

Buka alamat yang tampil, misalnya `http://192.168.1.20:8080`, dari komputer
lain di jaringan yang sama. Di Windows, izinkan Python di Windows Firewall
jika diminta.

**Pakai web server yang sudah ada (IIS, Apache/XAMPP, Nginx):** salin
seluruh isi folder ini (`index.html` + `vendor/`) ke folder web, misalnya
`C:\xampp\htdocs\rekap-faktur\`. Tidak perlu PHP, database, atau
konfigurasi lain.

## Kirim ke Google Sheets (opsional)

Setiap laporan menjadi 1 tab baru di Google Sheet, bernama tanggal laporan
(misal `25 September 2026`). Isi tab: rekap per penjual di kolom A–D dan
detail faktur di kolom F–I. Mengirim ulang laporan yang sama memperbarui
tab itu.

Pengaturan sekali saja:

1. Buka Google Sheet tujuan > **Extensions > Apps Script**.
2. Hapus isi `Code.gs`, tempel isi file `google-apps-script/Code.gs`, simpan.
3. **Project Settings** (ikon roda gigi) > **Script properties** > Add
   property: nama `REKAP_TOKEN`, nilai berupa kata sandi acak buatan Anda
   (misal `hnd-7Kq2x9Lm`). Simpan.
4. **Deploy > New deployment** > pilih tipe **Web app**:
   - Execute as: **Me**
   - Who has access: **Anyone**

   Klik Deploy, izinkan akses (Authorize), lalu salin **Web app URL**
   (berakhiran `/exec`).
5. Isi `config.js` di server:

   ```js
   window.REKAP_CONFIG = {
     sheetsUrl: "https://script.google.com/macros/s/XXXX/exec",
     sheetsToken: "hnd-7Kq2x9Lm",
   };
   ```

6. Muat ulang website. Tombol **Kirim ke Google Sheets** muncul di setiap
   hasil.

Jika `Code.gs` diubah nanti: **Deploy > Manage deployments > Edit > Version:
New version**, supaya URL tetap sama.

Keamanan: siapa pun yang bisa membuka website bisa melihat token di
`config.js`, jadi website ini sebaiknya hanya untuk jaringan kantor. Token
mencegah orang lain yang hanya tahu URL Apps Script menulis ke Sheet Anda.

## Menjalankan terus di server Linux (opsional)

Contoh service systemd `/etc/systemd/system/rekap-faktur.service`:

```
[Unit]
Description=Rekap Faktur Penjual
After=network.target

[Service]
WorkingDirectory=/opt/rekap-faktur
ExecStart=/usr/bin/python3 /opt/rekap-faktur/server.py 8080
Restart=always

[Install]
WantedBy=multi-user.target
```

Lalu: `sudo systemctl enable --now rekap-faktur`

## Catatan

- Kalau diakses dari internet (bukan jaringan kantor), pasang di belakang
  Nginx/IIS dengan HTTPS dan login.
- Font memakai Google Fonts bila ada internet; tanpa internet tampilan
  memakai font sistem, fungsi tetap sama.

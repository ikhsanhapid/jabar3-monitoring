# Jabar 3 Web Release

## Login Admin
- URL: `/admin`
- Username: `admin`
- Password default: `@Jabar3juara`

> Untuk hosting publik, sangat disarankan mengatur `ADMIN_PASS` dan `AUTH_SECRET` sebagai environment variable di server.

## Jalankan lokal
```bash
npm install
npm start
```
Buka `http://localhost:3000/` untuk dashboard dan `http://localhost:3000/admin` untuk admin.

## Upload Excel
Gunakan `.xlsx` atau `.xls`. Sheet pertama harus memiliki header wajib:
`KABUPATEN`, `KECAMATAN`, `PRODUK`, `AAE`, `PUD`, `ALOKASI 1 TAHUN`, `TOTAL S`.

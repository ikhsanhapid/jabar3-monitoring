# Jabar 3 Monitoring Web

Web dashboard Node.js untuk monitoring Jabar 3.

## Sumber data terpisah
1. **Data Penyaluran**: Excel `.xlsx/.xls` dengan kolom wajib `KABUPATEN, KECAMATAN, PRODUK, AAE, PUD, ALOKASI 1 TAHUN, TOTAL S` serta kolom bulanan.
2. **Data Stok**: Excel `.xlsx/.xls` terpisah. Header utama: `KABUPATEN, GUDANG LINI III, PPTS, PUD`.

## Admin
- URL: `/admin`
- Default username: `admin`
- Password: gunakan environment `ADMIN_PASS` (paket ini default `@Jabar3juara`, sebaiknya ganti untuk produksi).

## Environment
`ADMIN_USER`, `ADMIN_PASS`, `AUTH_SECRET`, `PORT`.

## Run
`npm install`
`node server.js`

Render: Build Command `npm install`, Start Command `node server.js`.

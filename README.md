# Oil Sounding Tank Calculator + History

## Fitur
- ST 1, ST 2, ST 3, ST 4 masing-masing membaca file Excel kalibrasinya.
- Tabel density + correction factor membaca `data/density.xlsx`.
- Perhitungan ullage, volume, density, mass, correction factor dan VMT.
- Setiap tank memiliki input sendiri.
- Tombol **Simpan ke History** menyimpan hasil sounding lengkap ke browser `localStorage`.
- Menu **History** dapat memfilter per tanggal, menghapus record, menghapus semua history, dan export CSV.

## Struktur data
```text
data/
├── ST 1.xlsx
├── ST 2.xlsx
├── ST 3.xlsx
├── ST 4.xlsx
└── density.xlsx
```

## Menjalankan
Gunakan VSCode + Live Server agar file XLSX dapat dibaca oleh browser.

## Catatan History
History tersimpan di browser/komputer yang digunakan. Jika browser storage dihapus atau aplikasi dibuka di komputer lain, history lokal tersebut tidak ikut berpindah. Gunakan **Export CSV** untuk backup.

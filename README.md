# Bunjir

Simulasi kapasitas drainase untuk enam kecamatan di Bandung Timur. Bunjir adalah proyek Berpikir Komputasional Kelompok 7. Namanya terinspirasi dari milk bun.

Bunjir membaca **21.888 jam data hujan NASA POWER**, menghitung debit limpasan dan kapasitas saluran, lalu menampilkan hasil dalam tabel terminal dan grafik PNG. Mode **Jam Rawan** menggabungkan jam berturut-turut saat limpasan melampaui kapasitas. Hasilnya menggambarkan potensi kelebihan kapasitas pada ruas atau skenario yang dianalisis.

## Cara menjalankan

Kloning repo, lalu pasang dependensi dan jalankan program dengan interpreter Python yang sama.

```powershell
git clone https://github.com/kumaru100/Bunjir.git
cd Bunjir
python -m pip install -r requirements.txt
python main.py
```

Jika editor memakai Python 3.13 pada Windows, gunakan `python3.13` untuk kedua perintah. Pilih **tahun → bulan → analisis**. Tahun 2024 dan 2025 tersedia Januari–Desember, sedangkan 2026 tersedia Januari–Juni. Tekan Enter di prompt path untuk memakai Excel yang disertakan.

Untuk melihat hasil Jam Rawan, pilih **2024 → 9 → 2 → Enter**. Panyileukan menghasilkan 6 jam potensi kelebihan kapasitas dan grafiknya tersimpan di `grafik/debit_2024-09_menu2.png`.

| Menu | Masukan | Hasil |
|---|---|---|
| 1. Analisis Bulanan | Pilih bulan dan file. Workbook bawaan memakai maksimum satu jam per bulan untuk keenam lokasi. | Debit, kapasitas, kelebihan debit, persentase pemakaian, status, dan cakupan data tiap lokasi. |
| 2. Analisis Jam Rawan | Pilih bulan dan file yang memiliki sheet `NASA per Jam`. | Waktu mulai, waktu selesai, durasi, dan puncak kelebihan debit setiap periode per lokasi. |

Keluaran terminal berbentuk tabel. Kedua menu juga menyimpan grafik PNG debit limpasan puncak dan kapasitas saluran tiap kecamatan ke folder `grafik/` di direktori tempat perintah dijalankan. Nama file memuat tahun, bulan, dan menu.

## Struktur Excel

Satu baris di sheet `Data` mewakili satu tahun, bulan, dan lokasi saluran. Untuk beberapa ruas di kecamatan yang sama, isi `Lokasi/Ruas` dengan nama yang berbeda.

| Kolom | Isi |
|---|---|
| `Tahun`, `Bulan`, `Kecamatan`, `Luas Tangkapan (ha)` | Tahun, nama bulan, kecamatan, dan luas yang mengalir ke ruas dalam hektare. Luas harus positif. Workbook lama tanpa `Tahun` dibaca sebagai 2026. |
| `Hujan (mm/jam)` | Angka nonnegatif untuk menu 1. Jika diisi rumus, simpan hasil kalkulasinya di Excel sebelum impor. Menu 2 membaca hujan dari sheet jam-jaman. |
| `Koefisien Limpasan C` | Koefisien langsung dari 0 sampai 1. |
| `Lebar (m)`, `Kedalaman Efektif (m)` | Ukuran positif penampang terbuka persegi panjang hingga batas yang dimodelkan. |
| `Kemiringan (%)`, `Manning n` | Nilai positif. Isi 0,50 sebagai angka untuk kemiringan 0,50%, sehingga S = 0,005 m/m. |
| `Jenis Data`, `Lokasi/Ruas`, `Sumber Data`, `Keterangan` | Status bukti, nama ruas, asal parameter, dan batasan. Jenis `AKTUAL` wajib dilengkapi lokasi serta sumber. |
| `Level Drainase`, `Material` | Kolom keterangan opsional. Perhitungan memakai lebar, kedalaman, dan n dari Excel. |

`AKTUAL` berarti penampang lokasi tertentu dilaporkan terukur. Parameter lain dalam baris itu tetap mengikuti jenis dan sumber masing-masing. `REFERENSI` adalah skenario berbasis sumber teknis, `ESTIMASI` adalah nilai turunan untuk lokasi, dan `ASUMSI` adalah nilai contoh tanpa pengukuran langsung. Kode sumber per parameter dijelaskan dalam sheet `Sumber Drainase` dan [REFERENSI.md](REFERENSI.md).

Workbook 2026 masih bisa dipakai karena sudah memuat C dan dimensi saluran. Jika file yang lebih lama hanya memuat level drainase dan fraksi kedap, pindahkan datanya ke format di atas sebelum impor.

## Rumus dan alur algoritma

1. Baca parameter lokasi dan validasi rentang masing-masing. Impor jam-jaman memeriksa tanggal WIB, menolak duplikat dan hujan negatif, lalu mengurutkan cap waktu.
2. Hitung debit limpasan dengan **Metode Rasional**. `Qh = C × i × A / 360` dalam m³/detik, dengan hujan `i` dalam mm/jam dan luas tangkapan `A` dalam ha. Untuk keluaran L/detik, program memakai `C × i × A × (10.000 / 3.600)`.
3. Hitung kapasitas saluran dengan **Manning** dan asumsi aliran seragam. Untuk lebar `b` dan kedalaman efektif `h` dalam meter, `As = b × h`, `R = As / (b + 2h)`, dan `Qd = As × R^(2/3) × sqrt(S) / n` dalam m³/detik. Program mengalikan Qd dengan 1.000 untuk mendapatkan L/detik.
4. Bandingkan debit. `Pemakaian = 100 × Qh / Qd` dan `Kelebihan debit = max(0, Qh − Qd)`. Status **AMAN** di bawah 80%, **WASPADA** pada 80–100%, dan **KAPASITAS TERLAMPAUI** di atas 100%. Ambang 80% adalah pilihan simulasi.
5. Untuk menu 2, lakukan perhitungan tiap jam. Jika jam yang melampaui kapasitas dimulai tepat saat periode sebelumnya berakhir, perpanjang periode dan simpan puncak kelebihan debit. Data hilang, hujan nol, atau debit di bawah atau tepat pada kapasitas memutus periode. Waktu selesai adalah batas akhir jam terakhir.

Alur itu menunjukkan dekomposisi (baca, validasi, hitung, kelompokkan), abstraksi (satu penampang per baris), pengenalan pola (jam berurutan), dan algoritma (perulangan serta percabangan). Kode memakai list untuk deret waktu dan dictionary untuk parameter lokasi.

## Cakupan dan interpretasi

- Hujan NASA POWER adalah `PRECTOTCORR` beresolusi satu jam untuk satu titik/grid bersama di −6,95° lintang dan 107,70° bujur. Tersedia **21.888 jam** dari Januari 2024 sampai Juni 2026 WIB, dengan timestamp UTC, WIB, dan maksimum per bulan. Intensitas satu jam merupakan pendekatan untuk durasi konsentrasi.
- Untuk 2026, Gedebage memakai satu ruas aktual, **A1–A2 Riung Arum Raya**. Untuk 2024 dan 2025, parameter studi 2026 dipakai sebagai skenario historis berlabel `ESTIMASI`. Lima kecamatan lain memakai luas asumsi, ukuran rujukan DAS Cinambo, C = 0,50 untuk skenario hunian multiunit terpisah, kemiringan 0,50%, dan n = 0,013 untuk skenario beton. Kondisi drainase aktual di lima kecamatan tersebut belum terverifikasi.
- Kelebihan debit adalah laju aliran dalam L/detik. Model menghasilkan potensi kelebihan kapasitas. Konfirmasi kejadian banjir, volume luapan aktual, dan tinggi genangan memerlukan data serta model tambahan. Model saat ini mengasumsikan saluran terbuka dengan aliran seragam dan analisis tiap ruas secara mandiri.

Untuk simulasi sedimentasi, ubah `Kedalaman Efektif (m)` di salinan workbook, beri label `ASUMSI` atau `ESTIMASI` yang sesuai, dan jalankan salah satu menu analisis dengan salinan itu.

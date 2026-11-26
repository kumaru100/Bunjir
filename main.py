# Sistem Monitoring Risiko Banjir Bandung Timur | Kelompok 7
from calendar import monthrange
from contextlib import closing
from datetime import datetime, timedelta
from math import isclose, isfinite, sqrt
from pathlib import Path

BULAN = ("Januari", "Februari", "Maret", "April", "Mei", "Juni",
         "Juli", "Agustus", "September", "Oktober", "November", "Desember")
JUMLAH_BULAN = {2024: 12, 2025: 12, 2026: 6}
SATU_JAM = timedelta(hours=1)
Hujan = "Hujan (mm/jam)"
Luas = "Luas Tangkapan (ha)"
C = "Koefisien Limpasan C"
Lebar, Dalam = "Lebar (m)", "Kedalaman Efektif (m)"
Kemiringan, Manning = "Kemiringan (%)", "Manning n"
Lokasi, Jenis = "Lokasi/Ruas", "Jenis Data"
KAPASITAS_TERLAMPAUI = "KAPASITAS TERLAMPAUI"


# Menghitung kecepatan, debit limpasan, dan kapasitas saluran.
def hitung_debit(saluran, hujan):
    # Metode Rasional dan Manning untuk penampang terbuka. Lihat REFERENSI.md.
    limpasan = saluran[C] * hujan * saluran[Luas] * (10000 / 3600)
    lebar, kedalaman = saluran[Lebar], saluran[Dalam]
    penampang = lebar * kedalaman
    radius = penampang / (lebar + 2 * kedalaman)
    kecepatan = radius ** (2 / 3) * sqrt(saluran[Kemiringan] / 100) / saluran[Manning]
    kapasitas = penampang * kecepatan * 1000
    if kapasitas <= 0:
        raise ValueError("Kapasitas saluran harus lebih dari nol.")
    return kecepatan, limpasan, kapasitas


# Menentukan status dari persentase kapasitas yang terpakai.
def status_debit(limpasan, kapasitas):
    pemakaian = limpasan / kapasitas
    # Abaikan selisih karena pembulatan floating point di kedua ambang.
    if pemakaian > 1 and not isclose(pemakaian, 1, rel_tol=1e-12):
        return KAPASITAS_TERLAMPAUI
    if pemakaian >= 0.8 or isclose(pemakaian, 0.8, rel_tol=1e-12):
        return "WASPADA"
    return "AMAN"


# Membaca angka dari Excel dan menunjukkan letak sel jika nilainya salah.
def angka(nilai, label):
    try:
        hasil = float(nilai)
        if not isinstance(nilai, bool) and isfinite(hasil):
            return hasil
    except (TypeError, ValueError):
        pass
    raise ValueError(f"{label} harus berisi angka hingga.")


# Membaca header dan memasangkan setiap sel dengan nama kolomnya.
def baris_excel(sheet, wajib):
    baris = sheet.iter_rows(values_only=True)
    header = [str(sel).strip() if sel is not None else "" for sel in next(baris, ())]
    salah = [nama for nama in wajib if header.count(nama) != 1]
    if salah:
        raise ValueError(f"Di '{sheet.title}', header ini harus ada tepat sekali: {', '.join(salah)}.")
    return enumerate((dict(zip(header, nilai)) for nilai in baris), 2)


# Mengambil parameter saluran untuk tahun dan bulan yang dipilih.
def baca_kecamatan(sheet, tahun, bulan, bulanan):
    positif = (Luas, Lebar, Dalam, Kemiringan, Manning)
    kolom_angka = (*positif, C, Hujan) if bulanan else (*positif, C)
    baris = baris_excel(sheet, ("Bulan", "Kecamatan", Jenis, *kolom_angka))
    data = {}
    for nomor, saluran in baris:
        if str(saluran["Bulan"]).strip().casefold() != bulan.casefold():
            continue
        if angka(saluran.get("Tahun", 2026), f"Tahun pada baris {nomor}") != tahun:
            continue
        nama = str(saluran["Kecamatan"] or "").strip()
        if not nama:
            raise ValueError(f"Nama kecamatan kosong pada baris {nomor}.")
        ruas = str(saluran.get(Lokasi) or "").strip()
        saluran["Kecamatan"], saluran[Lokasi] = nama, ruas
        jenis = str(saluran[Jenis]).strip().upper()
        if jenis not in {"AKTUAL", "REFERENSI", "ESTIMASI", "ASUMSI"}:
            raise ValueError(f"Jenis Data tidak valid pada baris {nomor}.")
        if jenis == "AKTUAL" and (not ruas or not str(saluran.get("Sumber Data") or "").strip()):
            raise ValueError(f"Data AKTUAL perlu lokasi dan sumber pada baris {nomor}.")
        saluran[Jenis] = jenis
        for kolom in kolom_angka:
            saluran[kolom] = angka(saluran[kolom], f"'{kolom}' pada baris {nomor}")
        if min(saluran[kolom] for kolom in positif) <= 0 or not 0 <= saluran[C] <= 1:
            raise ValueError(f"Baris {nomor}: luas, dimensi, kemiringan, dan n harus positif. C harus 0–1.")
        if bulanan and saluran[Hujan] < 0:
            raise ValueError(f"Curah hujan negatif pada baris {nomor}.")
        kunci = (nama.casefold(), ruas.casefold())
        if kunci in data:
            raise ValueError(f"Kecamatan/ruas duplikat pada baris {nomor}.")
        data[kunci] = saluran
    if not data:
        raise ValueError(f"Tidak ada data untuk {bulan} {tahun} di sheet '{sheet.title}'.")
    return data.values()


# Mengambil dan mengurutkan data hujan per jam untuk bulan terpilih.
def baca_hujan(sheet, tahun, bulan):
    waktu_wib, curah_hujan = "Waktu WIB", "PRECTOTCORR (mm/jam)"
    data = {}
    for nomor, row in baris_excel(sheet, (waktu_wib, curah_hujan)):
        waktu, hujan = row[waktu_wib], row[curah_hujan]
        if waktu is None and hujan is None:
            continue
        if not isinstance(waktu, datetime):
            raise ValueError(f"Waktu WIB pada baris {nomor} harus berupa tanggal dan jam Excel.")
        if (waktu.year, waktu.month) != (tahun, bulan):
            continue
        if waktu != waktu.replace(minute=0, second=0, microsecond=0) or waktu in data:
            raise ValueError(f"Waktu WIB pada baris {nomor} harus tepat pada awal jam dan unik.")
        hujan = None if hujan is None else angka(hujan, f"Hujan pada baris {nomor}")
        if hujan is not None and hujan != -999 and hujan < 0:
            raise ValueError(f"Curah hujan negatif pada baris {nomor}.")
        data[waktu] = hujan
    seri = sorted((waktu, hujan) for waktu, hujan in data.items() if hujan not in (None, -999))
    if not seri:
        raise ValueError(f"Tidak ada data hujan per jam untuk {BULAN[bulan - 1]} {tahun}.")
    return seri


# Menggabungkan jam berturut-turut saat limpasan melebihi kapasitas.
def cari_jam_rawan(saluran, seri):
    periode = []
    for waktu, hujan in seri:
        _, limpasan, kapasitas = hitung_debit(saluran, hujan)
        if status_debit(limpasan, kapasitas) != KAPASITAS_TERLAMPAUI:
            continue
        lebih = limpasan - kapasitas
        selesai = waktu + SATU_JAM
        # Jam hilang atau jam yang aman membuat waktu tidak menyambung.
        if periode and waktu == periode[-1][1]:
            mulai, _, puncak = periode[-1]
            periode[-1] = (mulai, selesai, max(puncak, lebih))
        else:
            periode.append((waktu, selesai, lebih))
    return periode


# Meminta angka lagi sampai masukan memenuhi syarat.
def minta_angka(pesan, konversi, valid):
    while True:
        try:
            nilai = konversi(input(pesan).replace(",", "."))
            if valid(nilai):
                return nilai
        except ValueError:
            pass
        print("Masukkan angka dalam rentang yang diminta.")


# Mencetak judul dan isi tabel dengan kolom yang rata.
def tabel(header, isi):
    baris = [header, *isi]
    lebar = [max(map(len, kolom)) for kolom in zip(*baris)]
    for i, row in enumerate(baris):
        print(" | ".join(nilai.ljust(panjang) for nilai, panjang in zip(row, lebar)))
        if i == 0:
            print("-+-".join("-" * panjang for panjang in lebar))


# Menampilkan ringkasan, detail jam rawan, dan menyiapkan data grafik.
def tampilkan(data, seri, tahun, bulan):
    puncak_hujan = None
    if seri is not None:
        jumlah_jam = monthrange(tahun, bulan)[1] * 24
        print(f"Data hujan WIB: {len(seri)}/{jumlah_jam} jam tersedia.")
        if len(seri) < jumlah_jam:
            print("Jam dengan data hilang memutus periode.")
        print("Waktu selesai adalah batas akhir jam terakhir.")
        puncak_hujan = max(hujan for _, hujan in seri)
    ringkasan, parameter, detail, grafik = [], [], [], []
    nama_daerah = [saluran["Kecamatan"] for saluran in data]
    for nomor, saluran in enumerate(data, 1):
        nama = saluran["Kecamatan"]
        label = f"{nomor}. {nama}" if nama_daerah.count(nama) > 1 else nama
        hujan = saluran[Hujan] if seri is None else puncak_hujan
        v, q, kapasitas = hitung_debit(saluran, hujan)
        grafik.append((label, q, kapasitas))
        row = [str(nomor), nama, saluran[Jenis], f"{hujan:.2f}", f"{q:.2f}",
               f"{kapasitas:.2f}", f"{max(0, q - kapasitas):.2f}", f"{q / kapasitas:.1%}",
               status_debit(q, kapasitas)]
        parameter.append([str(nomor), saluran[Lokasi] or "-", f"{saluran[Luas]:g}", f"{saluran[C]:g}",
                          f"{saluran[Lebar]:g} × {saluran[Dalam]:g}", f"{v:.2f}"])
        if seri is not None:
            periode = cari_jam_rawan(saluran, seri)
            row.append(str(sum((akhir - awal) // SATU_JAM for awal, akhir, _ in periode)))
            for mulai, selesai, puncak in periode:
                detail.append([str(nomor), nama, f"{mulai:%d/%m %H:%M}", f"{selesai:%d/%m %H:%M}",
                               str((selesai - mulai) // SATU_JAM), f"{puncak:.2f}"])
        ringkasan.append(row)
    print("\nDebit dalam L/dtk. Ringkasan memakai intensitas puncak untuk menu Jam Rawan.")
    header = ["No", "Kecamatan", "Data", "mm/jam", "Limpasan", "Kapasitas", "Lebih", "Pakai", "Status"]
    tabel(header + (["Jam rawan"] if seri is not None else []), ringkasan)
    if detail:
        print("\nPeriode potensi kelebihan kapasitas")
        tabel(["No", "Kecamatan", "Mulai WIB", "Selesai WIB", "Jam", "Lebih maks."], detail)
    elif seri is not None:
        print("\nTidak ada jam rawan pada data yang tersedia.")
    print("\nRuas dan cakupan data")
    tabel(["No", "Ruas", "A (ha)", "C", "b × h (m)", "V (m/dtk)"], parameter)
    for nomor, saluran in enumerate(data, 1):
        catatan = saluran.get("Keterangan") or ""
        if saluran[Jenis] != "AKTUAL" and "belum terverifikasi" not in catatan:
            catatan = f"Kondisi drainase aktual belum terverifikasi. {catatan}".strip()
        if catatan:
            print(f"{nomor}. {catatan}")
    print("Hasil simulasi. Ambang WASPADA: 80–100%.")
    return grafik


# Menyimpan grafik perbandingan limpasan dan kapasitas sebagai PNG.
def simpan_grafik(data, tahun, bulan, pilihan):
    import matplotlib

    matplotlib.use("Agg")
    from matplotlib import pyplot as plt

    nama, limpasan, kapasitas = zip(*data)
    posisi = range(len(nama))
    fig, ax = plt.subplots(figsize=(10, max(4, 0.60 * len(nama) + 1.7)), layout="constrained")
    ax.barh([i - 0.17 for i in posisi], limpasan, height=0.32, color="#176A95", label="Debit limpasan puncak")
    ax.barh([i + 0.17 for i in posisi], kapasitas, height=0.32, color="#9AAFC0", label="Kapasitas saluran")
    ax.set_yticks(posisi, nama)
    ax.invert_yaxis()
    ax.set_xlabel("Debit (L/detik)")
    judul = f"{BULAN[bulan - 1]} {tahun}"
    ax.set_title(f"Perbandingan debit antar kecamatan\n{judul} · Simulasi", loc="left")
    ax.set_axisbelow(True)
    ax.grid(axis="x", alpha=0.24)
    ax.legend(loc="upper center", bbox_to_anchor=(0.5, -0.14), ncols=2, frameon=False)
    folder = Path.cwd() / "grafik"
    folder.mkdir(exist_ok=True)
    nama_menu = f"menu{pilihan}" if isinstance(pilihan, int) else pilihan
    file = folder / f"debit_{tahun}-{bulan:02d}_{nama_menu}.png"
    fig.savefig(file, dpi=160)
    plt.close(fig)
    print(f"Grafik: {file}")


# Membuka Excel dan membaca parameter saluran serta hujan sesuai menu.
def impor_excel(pilihan, tahun, bulan):
    from openpyxl import load_workbook

    bawaan = Path(__file__).with_name("Data_Bandung_Timur_NASA_2024_2026.xlsx")
    path = Path(input(f"Path Excel (Enter untuk {bawaan.name}): ").strip().strip('"') or bawaan)
    if path.suffix.lower() != ".xlsx":
        raise ValueError("Gunakan file Excel .xlsx.")
    with closing(load_workbook(path, read_only=True, data_only=True)) as wb:
        sheet = wb["Data"] if "Data" in wb.sheetnames else wb.active
        data = list(baca_kecamatan(sheet, tahun, BULAN[bulan - 1], bulanan=pilihan == 1))
        if pilihan == 2 and "NASA per Jam" not in wb.sheetnames:
            raise ValueError("Jam Rawan memerlukan sheet 'NASA per Jam'.")
        seri = baca_hujan(wb["NASA per Jam"], tahun, bulan) if pilihan == 2 else None
    return data, seri


# Menjalankan pilihan tahun, bulan, dan menu analisis.
def main():
    print("=== MONITORING RISIKO BANJIR BANDUNG TIMUR ===")
    print("1. Analisis data Bandung\n2. Bikin simulasi manual")
    mode = minta_angka("Pilih mode (1–2): ", int, lambda nilai: 1 <= nilai <= 2)
    if mode == 1:
        tahun = minta_angka("Pilih tahun (2024/2025/2026): ", int,
                            lambda nilai: nilai in JUMLAH_BULAN)
        jumlah_bulan = JUMLAH_BULAN[tahun]
        for i, nama in enumerate(BULAN[:jumlah_bulan], 1):
            print(f"{i}. {nama}")
        bulan = minta_angka(f"Pilih bulan (1–{jumlah_bulan}): ", int,
                            lambda nilai: 1 <= nilai <= jumlah_bulan)
        print("1. Analisis Bulanan\n2. Analisis Harian untuk Jam Rawan")
        pilihan = minta_angka("Pilih analisis (1–2): ", int, lambda nilai: 1 <= nilai <= 2)
        data, seri = impor_excel(pilihan, tahun, bulan)
        nama_grafik = pilihan
    else:
        tahun = minta_angka("Tahun simulasi: ", int, lambda nilai: nilai > 0)
        bulan = minta_angka("Bulan simulasi (1–12): ", int, lambda nilai: 1 <= nilai <= 12)
        jumlah_daerah = minta_angka("Jumlah daerah: ", int, lambda nilai: nilai > 0)
        data, seri = [], None
        for nomor in range(1, jumlah_daerah + 1):
            print(f"\nDaerah {nomor} dari {jumlah_daerah}")
            nama = input("Nama daerah: ").strip() or f"Daerah {nomor}"
            data.append({
                "Kecamatan": nama,
                Lokasi: "Input manual",
                Jenis: "ASUMSI",
                Luas: minta_angka("Luas tangkapan (ha): ", float,
                                  lambda nilai: isfinite(nilai) and nilai > 0),
                Hujan: minta_angka(f"Intensitas hujan {BULAN[bulan - 1]} (mm/jam): ", float,
                                   lambda nilai: isfinite(nilai) and nilai >= 0),
                C: minta_angka("Koefisien limpasan C (0–1): ", float,
                              lambda nilai: isfinite(nilai) and 0 <= nilai <= 1),
                Lebar: minta_angka("Lebar saluran (m): ", float,
                                   lambda nilai: isfinite(nilai) and nilai > 0),
                Dalam: minta_angka("Kedalaman saluran (m): ", float,
                                   lambda nilai: isfinite(nilai) and nilai > 0),
                Kemiringan: minta_angka("Kemiringan saluran (%): ", float,
                                        lambda nilai: isfinite(nilai) and nilai > 0),
                Manning: minta_angka("Manning n: ", float,
                                    lambda nilai: isfinite(nilai) and nilai > 0),
                "Keterangan": "Parameter simulasi diisi manual oleh pengguna.",
            })
        nama_grafik = "simulasi"
    print(f"\nBulan: {BULAN[bulan - 1]} {tahun}")
    grafik = tampilkan(data, seri, tahun, bulan)
    simpan_grafik(grafik, tahun, bulan, nama_grafik)


if __name__ == "__main__":
    main()

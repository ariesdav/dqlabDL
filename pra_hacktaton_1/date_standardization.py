import pandas as pd

def normalize_tanggal_transaksi(input_xlsx_path: str, output_xlsx_path: str) -> None:
    # Baca file dan sheet "transaksi"
    df = pd.read_excel(input_xlsx_path, sheet_name="transaksi", dtype=str)
   
    # Bersihkan tanda baca
    df['Tanggal Transaksi'] = (
        df['Tanggal Transaksi']
        .astype(str)
        .str.replace(",", "", regex=False)
        .str.replace("'", "", regex=False)
        .str.strip()
    )

    # Ganti nama bulan ke angka
    bulan = {
        "Januari": "01", "January": "01", "Jan": "01",
        "Februari": "02", "February": "02", "Feb": "02",
        "Maret": "03", "March": "03", "Mar": "03",
        "April": "04", "Apr": "04",
        "Mei": "05", "May": "05",
        "Juni": "06", "June": "06", "Jun": "06",
        "Juli": "07", "July": "07", "Jul": "07",
        "Agustus": "08", "August": "08", "Aug": "08", "Agu": "08",
        "September": "09", "Sep": "09", "Sept": "09",
        "Oktober": "10", "October": "10", "Okt": "10", "Oct": "10",
        "November": "11", "Nov": "11",
        "Desember": "12", "December": "12", "Des": "12", "Dec": "12"
    }
    for nama, angka in bulan.items():
        df['Tanggal Transaksi'] = df['Tanggal Transaksi'].str.replace(nama, angka, regex=False)

    # Format “2024, 27 08” ke “27 08 2024”
    df['Tanggal Transaksi'] = df['Tanggal Transaksi'].apply(
        lambda x: " ".join(x.split()[1:] + x.split()[:1])
        if x.split()[0].isdigit() and len(x.split()[0]) == 4 else x
    )

    # Tambah "20" ke tahun 2 digit di akhir
    df['Tanggal Transaksi'] = df['Tanggal Transaksi'].str.replace(
        r"(\s)(\d{2})$", r"\g<1>20\2", regex=True
    )

    # Ubah ke format tanggal
    df['Tanggal Transaksi'] = pd.to_datetime(
        df['Tanggal Transaksi'], dayfirst=True, errors='coerce'
    ).dt.strftime('%d-%m-%Y')

    # Simpan ke file baru
    with pd.ExcelWriter(output_xlsx_path, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name="transaksi", index=False)

# Jalankan fungsi
normalize_tanggal_transaksi("penjualan_dqmart_01-beta.xlsx", "penjualan_dqmart_01-output.xlsx")
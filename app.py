import streamlit as st
import pymysql
import pandas as pd

st.title("Aplikasi Database Pegawai")

# Pilih tabel
tabel = st.selectbox(
    "Pilih Tabel:",
    ["mstpegawai", "fgaji"]
)

# Filter khusus mstpegawai
status = "Semua"
if tabel == "mstpegawai":
    status = st.selectbox(
        "Status Pegawai:",
        ["Semua", "Aktif", "Tidak Aktif"]
    )

# Batas jumlah baris khusus fgaji
limit = 1000
cari_berdasarkan = "NIP"
keyword = ""
checked_cols = []
if tabel == "fgaji":
    limit = st.number_input("Jumlah Baris Data:", min_value=100, max_value=100000, value=1000, step=100)
    perhitungan_fgaji = st.selectbox("Pilih Perhitungan:", ["Hitung Kelebihan Bayar Gaji", "Hitung Masa Kerja"])
    cari_berdasarkan = "NIP"
    keyword = ""
    if perhitungan_fgaji == "Hitung Kelebihan Bayar Gaji":
        cari_berdasarkan = st.selectbox("Cari Berdasarkan:", ["NIP", "NAMA"])
        keyword = st.text_input(f"Masukkan {cari_berdasarkan}:")
        st.write("Hitung Kelebihan Bayar Gaji:")
        tgl_gaji_mulai = st.date_input(
            "TGLGAJI (Mulai Pengembalian):",
            min_value=pd.Timestamp('2014-01-01').date(),
            max_value=pd.Timestamp('2035-12-31').date()
        )
        tgl_gaji_sampai = st.date_input(
            "TGLGAJI (Sampai Pengembalian):",
            min_value=pd.Timestamp('2014-01-01').date(),
            max_value=pd.Timestamp('2035-12-31').date()
        )
        if st.checkbox("Gapok"): checked_cols.append("GAPOK")
        if st.checkbox("Suami/Istri"): checked_cols.append("TJISTRI")
        if st.checkbox("Anak"): checked_cols.append("TJANAK")
        if st.checkbox("Tunjangan Struktural"): checked_cols.append("TJESELON")
        if st.checkbox("Tunjangan Fungsional"): checked_cols.append("TJFUNGSI")
        if st.checkbox("Tunjangan Khusus"): checked_cols.append("TJKHUSUS")
        if st.checkbox("Tunjangan Beras"): checked_cols.append("TJBERAS")
        if st.checkbox("Tunjangan Umum"): checked_cols.append("TJUMUM")
        if st.checkbox("Pembulatan"): checked_cols.append("TBULAT")    
        if st.checkbox("Gaji Bersih"): checked_cols.append("BERSIH")

try:
    conn = pymysql.connect(
        host="localhost",
        port=3000,
        user="taspen",
        password="taspen",
        database="meifinal2026"
    )

    cursor = conn.cursor()

    # Query
    if tabel == "mstpegawai":
        if status == "Aktif":
            query = "SELECT * FROM mstpegawai WHERE TMTSTOP IS NULL"
        elif status == "Tidak Aktif":
            query = "SELECT * FROM mstpegawai WHERE TMTSTOP IS NOT NULL"
        else:
            query = "SELECT * FROM mstpegawai"
    else:
        where_clause = ""
        if keyword:
            kolom = "NIP" if cari_berdasarkan == "NIP" else "NAMA"
            where_clause = f" WHERE {kolom} LIKE '%{keyword}%'"
        query = f"SELECT * FROM fgaji{where_clause} LIMIT {limit}"

    cursor.execute(query)

    rows = cursor.fetchall()
    columns = [col[0] for col in cursor.description]

    df = pd.DataFrame(rows, columns=columns)

    # Hitung kelebihan bayar gaji untuk fgaji
    if tabel == "fgaji" and perhitungan_fgaji == "Hitung Kelebihan Bayar Gaji":
        # Filter data fgaji sesuai range TGLGAJI yang dipilih
        if 'tgl_gaji_mulai' in locals() and 'tgl_gaji_sampai' in locals() and 'TGLGAJI' in df.columns:
            df['TGLGAJI_temp'] = pd.to_datetime(df['TGLGAJI'], errors='coerce').dt.date
            df = df[(df['TGLGAJI_temp'] >= tgl_gaji_mulai) & (df['TGLGAJI_temp'] <= tgl_gaji_sampai)].copy()
            df = df.drop(columns=['TGLGAJI_temp'])

        if checked_cols:
            for col in checked_cols:
                if col in df.columns:
                    df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
            df['Kelebihan Bayar'] = df[checked_cols].sum(axis=1)

            # Tampilkan hanya kolom dasar + yang dichecklist + Kelebihan Bayar
            kolom_dasar = ['TGLGAJI', 'NIP', 'NAMA']
            kolom_tampilan = kolom_dasar + checked_cols + ['Kelebihan Bayar']
            kolom_ada = [col for col in kolom_tampilan if col in df.columns]
            df = df[kolom_ada]

            # Buat baris total untuk kolom numerik
            total_data = {col: '' for col in df.columns}
            total_data['NIP'] = 'TOTAL'
            total_data['Kelebihan Bayar'] = df['Kelebihan Bayar'].sum()
            for col in checked_cols:
                if col in df.columns:
                    total_data[col] = df[col].sum()
            df = pd.concat([df, pd.DataFrame([total_data])], ignore_index=True)

            # Simpan df untuk export (sebelum format string)
            df_export = df.copy()

            # Format angka dengan titik pemisah ribuan
            for col in checked_cols + ['Kelebihan Bayar']:
                if col in df.columns:
                    df[col] = df[col].apply(lambda x: f"{x:,.0f}".replace(",", ".") if isinstance(x, (int, float)) else x)

    if tabel == "fgaji":
        st.write(f"Tabel: {tabel} | Perhitungan: {perhitungan_fgaji}")
    else:
        st.write(f"Tabel: {tabel} | Status: {status}")
    st.dataframe(df)

    # Tombol Export ke Excel
    if tabel == "fgaji" and checked_cols and perhitungan_fgaji == "Hitung Kelebihan Bayar Gaji":
        from io import BytesIO
        
        def generate_excel(df_to_export, format_cols):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_to_export.to_excel(writer, index=False, sheet_name='Kelebihan Bayar')
                workbook = writer.book
                worksheet = writer.sheets['Kelebihan Bayar']
                
                # Format angka dengan titik pemisah ribuan & Auto-fit lebar kolom
                for idx, col_name in enumerate(df_to_export.columns):
                    col_letter = worksheet.cell(row=1, column=idx+1).column_letter
                    max_length = len(str(col_name))
                    
                    for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row, min_col=idx+1, max_col=idx+1):
                        for cell in row:
                            try:
                                if cell.value is not None and len(str(cell.value)) > max_length:
                                    max_length = len(str(cell.value))
                            except:
                                pass
                            
                            # Format angka dengan pemisah ribuan (titik)
                            if col_name in format_cols and isinstance(cell.value, (int, float)):
                                cell.number_format = '#.##0'
                                
                    adjusted_width = max_length + 2
                    worksheet.column_dimensions[col_letter].width = adjusted_width
                    
            return output.getvalue()

        st.download_button(
            label="📥 Export ke Excel (XLSX)",
            data=lambda: generate_excel(df_export, checked_cols + ['Kelebihan Bayar']),
            file_name='kelebihan_bayar.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

    conn.close()

except Exception as e:
    st.error(f"ERROR: {e}")
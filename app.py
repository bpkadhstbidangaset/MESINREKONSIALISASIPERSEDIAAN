import io
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Rekonsiliasi Persediaan vs LRA", layout="wide"
)
st.title("Aplikasi Rekonsiliasi Persediaan vs LRA (SIPD)")
st.caption("Pemerintah Kabupaten Hulu Sungai Tengah")

# Sidebar Upload File
st.sidebar.header("Upload File Excel")
file_rak = st.sidebar.file_uploader(
    "1. RAK PERSEDIAAN.xlsx", type=["xlsx"]
)
file_lra = st.sidebar.file_uploader(
    "2. DATA REALISASI (LRA/SIPD).xlsx", type=["xlsx"]
)
file_app = st.sidebar.file_uploader(
    "3. DATA APLIKASI PERSEDIAAN.xlsx", type=["xlsx"]
)

if file_rak and file_lra and file_app:
    try:
        # 1. BACA MASTER RAK PERSEDIAAN
        df_rak = pd.read_excel(file_rak, skiprows=1)
        rak_codes = set(
            df_rak[
                "KLASIFIKASI AKUN REK PERSEDIAAN (Sebagai Referensi Pencatatan Persedian)"
            ]
            .dropna()
            .astype(str)
            .str.strip()
        )
        rak_codes = {c for c in rak_codes if c.startswith("5.")}

        # 2. BACA DATA LRA / SIPD
        xls_lra = pd.ExcelFile(file_lra)
        excl_keywords = []
        if "persediaan" in xls_lra.sheet_names:
            df_excl = pd.read_excel(file_lra, sheet_name="persediaan")
            excl_keywords = (
                df_excl.iloc[:, 1].dropna().astype(str).str.lower().tolist()
            )

        sheet_lra_name = (
            "01" if "01" in xls_lra.sheet_names else xls_lra.sheet_names[0]
        )
        raw_lra = pd.read_excel(file_lra, sheet_name=sheet_lra_name)

        header_idx = 3
        for idx, row in raw_lra.iterrows():
            if any(row.astype(str).str.contains("Kode Rekening", na=False)):
                header_idx = idx
                break

        df_lra = raw_lra.iloc[header_idx + 1 :].copy()
        df_lra.columns = raw_lra.iloc[header_idx].values
        df_lra["Kode_Clean"] = (
            df_lra["Kode Rekening"].astype(str).str.strip()
        )

        # FILTER HANYA KODE YANG TERDAFTAR DI RAK PERSEDIAAN
        df_lra_persediaan = df_lra[df_lra["Kode_Clean"].isin(rak_codes)].copy()

        # EKSKLUSI PUSAT (seperti Pakaian/Makan Minum Rapat jika ada)
        for kw in excl_keywords:
            df_lra_persediaan = df_lra_persediaan[
                ~df_lra_persediaan["Nama Rekening"]
                .astype(str)
                .str.lower()
                .str.contains(kw)
            ]

        # FILTER SKPD
        st.sidebar.markdown("---")
        st.sidebar.header("Filter SKPD")
        list_skpd = sorted(
            df_lra_persediaan["Nama SKPD"].dropna().unique().tolist()
        )
        skpd_pilihan = st.sidebar.selectbox(
            "Pilih SKPD yang ingin direkonsiliasi:", list_skpd
        )

        df_lra_filtered = df_lra_persediaan[
            df_lra_persediaan["Nama SKPD"] == skpd_pilihan
        ].copy()

        # REKAP LRA
        lra_summary = (
            df_lra_filtered.groupby(["Kode_Clean", "Nama Rekening"])[
                "Nilai Realisasi"
            ]
            .sum()
            .reset_index()
        )
        lra_summary.rename(
            columns={
                "Kode_Clean": "Kode_Rekening",
                "Nilai Realisasi": "Total_LRA",
            },
            inplace=True,
        )

        # 3. BACA DATA APLIKASI PERSEDIAAN
        df_app = pd.read_excel(file_app, skiprows=9)
        cols = [
            "No",
            "Tanggal",
            "Referensi",
            "No_Pesanan",
            "Sumber_Dana",
            "Kode_Belanja",
            "Nama_Belanja",
            "Nilai_Belanja",
            "Kode_Persediaan",
            "Nama_Persediaan",
            "Nilai_Persediaan",
        ]
        df_app.columns = cols[: len(df_app.columns)]

        meta_cols = [
            "No",
            "Tanggal",
            "Referensi",
            "No_Pesanan",
            "Sumber_Dana",
            "Kode_Belanja",
            "Nama_Belanja",
            "Nilai_Belanja",
        ]
        df_app[meta_cols] = df_app[meta_cols].ffill()
        df_app = df_app[df_app["Kode_Belanja"] != "Kode"].dropna(
            subset=["Kode_Persediaan"]
        )

        df_app["Kode_Rekening"] = (
            df_app["Kode_Belanja"].astype(str).str.strip()
        )
        app_summary = (
            df_app.groupby("Kode_Rekening")["Nilai_Persediaan"]
            .sum()
            .reset_index()
        )
        app_summary.rename(
            columns={"Nilai_Persediaan": "Total_Persediaan"}, inplace=True
        )

        # 4. GABUNGKAN DATA (REKONSILIASI)
        recon = pd.merge(
            lra_summary, app_summary, on="Kode_Rekening", how="outer"
        ).fillna(0)
        recon["Selisih"] = recon["Total_LRA"] - recon["Total_Persediaan"]

        def status_rekonsiliasi(row):
            if row["Selisih"] == 0:
                return "Cocok / Balance"
            elif row["Total_Persediaan"] == 0:
                return "Belum Input di Aplikasi Persediaan"
            elif row["Total_LRA"] == 0:
                return "Tidak Ada Realisasi di LRA"
            else:
                return "Beda Nilai Realisasi"

        recon["Status"] = recon.apply(status_rekonsiliasi, axis=1)

        # DASHBOARD
        st.subheader(f"Hasil Rekonsiliasi SKPD: {skpd_pilihan}")

        col1, col2, col3 = st.columns(3)
        col1.metric("Total LRA", f"Rp {recon['Total_LRA'].sum():,.0f}")
        col2.metric(
            "Total Aplikasi Persediaan",
            f"Rp {recon['Total_Persediaan'].sum():,.0f}",
        )
        col3.metric(
            "Total Selisih", f"Rp {abs(recon['Selisih'].sum()):,.0f}"
        )

        st.markdown("---")

        recon_display = recon.copy()
        recon_display["Total_LRA"] = recon_display["Total_LRA"].map(
            "Rp {:,.0f}".format
        )
        recon_display["Total_Persediaan"] = recon_display[
            "Total_Persediaan"
        ].map("Rp {:,.0f}".format)
        recon_display["Selisih"] = recon_display["Selisih"].map(
            "Rp {:,.0f}".format
        )

        st.subheader("Detail Rekonsiliasi Per Kode Rekening")
        st.dataframe(recon_display, use_container_width=True)

        # DOWNLOAD EXCEL
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            recon.to_excel(
                writer, sheet_name="Hasil Rekonsiliasi", index=False
            )

        st.download_button(
            label="📥 Download Hasil Rekonsiliasi (.xlsx)",
            data=buffer.getvalue(),
            file_name=f"Hasil_Rekonsiliasi_{skpd_pilihan}.xlsx",
            mime="application/vnd.ms-excel",
        )

    except Exception as e:
        st.error(f"Terjadi kesalahan saat memproses data: {e}")

else:
    st.info("Silakan unggah ketiga file Excel di menu sebelah kiri (Sidebar).")

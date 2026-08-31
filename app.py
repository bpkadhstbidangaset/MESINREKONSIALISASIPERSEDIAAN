import io
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Rekonsiliasi Persediaan vs LRA",
    page_icon="📊",
    layout="wide",
)

st.title("📊 Aplikasi Rekonsiliasi Persediaan vs LRA (SIPD)")
st.caption("Pemerintah Kabupaten Hulu Sungai Tengah")

# Helper function untuk format Rupiah Indonesia
def format_rupiah(val):
    if pd.isna(val):
        return "Rp 0"
    return f"Rp {val:,.0f}".replace(",", "_").replace(".", ",").replace("_", ".")

# Helper function untuk pembersihan angka
def clean_numeric(series):
    return (
        series.astype(str)
        .str.replace("Rp", "", regex=False)
        .str.replace(" ", "", regex=False)
        .str.replace(".", "", regex=False)
        .str.replace(",", ".", regex=False)
        .replace(["nan", "None", ""], 0)
        .astype(float)
    )

# ================= SIDEBAR: UPLOAD FILE =================
st.sidebar.header("📁 Upload File Excel")
file_rak = st.sidebar.file_uploader("1. Master RAK Persediaan (.xlsx)", type=["xlsx"])
file_lra = st.sidebar.file_uploader("2. Data Realisasi LRA/SIPD (.xlsx)", type=["xlsx"])
file_app = st.sidebar.file_uploader("3. Data Aplikasi Persediaan (.xlsx)", type=["xlsx"])

if file_rak and file_lra and file_app:
    try:
        with st.spinner("Memproses dan membaca berkas Excel..."):
            # ----------------------------------------------------
            # 1. BACA MASTER RAK PERSEDIAAN
            # ----------------------------------------------------
            df_rak = pd.read_excel(file_rak, skiprows=1)
            first_col_rak = df_rak.columns[0]
            rak_codes = set(
                df_rak[first_col_rak]
                .dropna()
                .astype(str)
                .str.strip()
            )
            rak_codes = {c for c in rak_codes if c.startswith("5.")}

            # ----------------------------------------------------
            # 2. BACA DATA LRA / SIPD
            # ----------------------------------------------------
            xls_lra = pd.ExcelFile(file_lra)
            excl_keywords = []
            if "persediaan" in [s.lower() for s in xls_lra.sheet_names]:
                sheet_excl_name = [s for s in xls_lra.sheet_names if s.lower() == "persediaan"][0]
                df_excl = pd.read_excel(file_lra, sheet_name=sheet_excl_name)
                excl_keywords = (
                    df_excl.iloc[:, 1].dropna().astype(str).str.lower().tolist()
                )

            sheet_lra_name = "01" if "01" in xls_lra.sheet_names else xls_lra.sheet_names[0]
            raw_lra = pd.read_excel(file_lra, sheet_name=sheet_lra_name)

            # Cari baris header secara otomatis
            header_idx = 3
            for idx, row in raw_lra.iterrows():
                if any(row.astype(str).str.contains("Kode Rekening", case=False, na=False)):
                    header_idx = idx
                    break

            df_lra = raw_lra.iloc[header_idx + 1 :].copy()
            df_lra.columns = [str(col).strip() for col in raw_lra.iloc[header_idx].values]
            df_lra["Kode_Clean"] = df_lra["Kode Rekening"].astype(str).str.strip()

            # Filter akun persediaan berdasarkan RAK
            df_lra_persediaan = df_lra[df_lra["Kode_Clean"].isin(rak_codes)].copy()

            # Filter pengecualian kata kunci
            for kw in excl_keywords:
                df_lra_persediaan = df_lra_persediaan[
                    ~df_lra_persediaan["Nama Rekening"]
                    .astype(str)
                    .str.lower()
                    .str.contains(kw)
                ]

            # Konversi Nilai Realisasi ke numerik
            df_lra_persediaan["Nilai Realisasi"] = pd.to_numeric(
                df_lra_persediaan["Nilai Realisasi"], errors="coerce"
            ).fillna(0)

            # ----------------------------------------------------
            # 3. BACA DATA APLIKASI PERSEDIAAN
            # ----------------------------------------------------
            df_app = pd.read_excel(file_app, skiprows=9)
            cols = [
                "No", "Tanggal", "Referensi", "No_Pesanan", "Sumber_Dana",
                "Kode_Belanja", "Nama_Belanja", "Nilai_Belanja",
                "Kode_Persediaan", "Nama_Persediaan", "Nilai_Persediaan"
            ]
            df_app.columns = cols[: len(df_app.columns)]

            meta_cols = [
                "No", "Tanggal", "Referensi", "No_Pesanan", "Sumber_Dana",
                "Kode_Belanja", "Nama_Belanja", "Nilai_Belanja"
            ]
            df_app[meta_cols] = df_app[meta_cols].ffill()
            df_app = df_app[df_app["Kode_Belanja"] != "Kode"].dropna(subset=["Kode_Persediaan"]).copy()

            # Parsing Tanggal & Nilai Persediaan
            df_app["Tanggal_Parsed"] = pd.to_datetime(df_app["Tanggal"], errors="coerce")
            df_app["Bulan"] = df_app["Tanggal_Parsed"].dt.month
            df_app["Bulan_Nama"] = df_app["Tanggal_Parsed"].dt.strftime("%B")
            df_app["Kode_Rekening"] = df_app["Kode_Belanja"].astype(str).str.strip()
            df_app["Nilai_Persediaan"] = pd.to_numeric(df_app["Nilai_Persediaan"], errors="coerce").fillna(0)

        # ================= SIDEBAR FILTER =================
        st.sidebar.markdown("---")
        st.sidebar.subheader("⚙️ Filter Rekonsiliasi")

        # 1. Filter SKPD
        list_skpd = sorted(df_lra_persediaan["Nama SKPD"].dropna().unique().tolist())
        if not list_skpd:
            st.warning("Tidak ditemukan data SKPD di file LRA.")
            st.stop()

        skpd_pilihan = st.sidebar.selectbox("Pilih SKPD:", list_skpd)

        # 2. Filter Periode / Bulan
        filter_mode = st.sidebar.radio(
            "Mode Periode Rekonsiliasi:",
            ["Semua Periode (YTD)", "Per Bulan", "Rentang Tanggal Custom"]
        )

        df_app_filtered = df_app.copy()

        nama_bulan_id = {
            1: "Januari", 2: "Februari", 3: "Maret", 4: "April",
            5: "Mei", 6: "Juni", 7: "Juli", 8: "Agustus",
            9: "September", 10: "Oktober", 11: "November", 12: "Desember"
        }

        periode_label = "Semua Periode"

        if filter_mode == "Per Bulan":
            bulan_pilihan = st.sidebar.selectbox(
                "Pilih Bulan:",
                options=list(nama_bulan_id.keys()),
                format_func=lambda x: f"{nama_bulan_id[x]} (Bulan {x})"
            )
            df_app_filtered = df_app_filtered[df_app_filtered["Bulan"] == bulan_pilihan]
            periode_label = f"Bulan {nama_bulan_id[bulan_pilihan]}"

        elif filter_mode == "Rentang Tanggal Custom":
            min_date = df_app["Tanggal_Parsed"].min()
            max_date = df_app["Tanggal_Parsed"].max()
            date_range = st.sidebar.date_input(
                "Pilih Rentang Tanggal:",
                value=(min_date if pd.notna(min_date) else None, max_date if pd.notna(max_date) else None)
            )
            if isinstance(date_range, (list, tuple)) and len(date_range) == 2:
                tgl_awal, tgl_akhir = date_range
                df_app_filtered = df_app_filtered[
                    (df_app_filtered["Tanggal_Parsed"].dt.date >= tgl_awal) &
                    (df_app_filtered["Tanggal_Parsed"].dt.date <= tgl_akhir)
                ]
                periode_label = f"{tgl_awal} s.d. {tgl_akhir}"

        # ----------------------------------------------------
        # 4. AGREGASI & MERGE
        # ----------------------------------------------------
        df_lra_filtered = df_lra_persediaan[df_lra_persediaan["Nama SKPD"] == skpd_pilihan].copy()

        lra_summary = (
            df_lra_filtered.groupby(["Kode_Clean", "Nama Rekening"])[
                "Nilai Realisasi"
            ]
            .sum()
            .reset_index()
            .rename(
                columns={
                    "Kode_Clean": "Kode_Rekening",
                    "Nama Rekening": "Nama_Rekening_LRA",
                    "Nilai Realisasi": "Total_LRA",
                }
            )
        )

        app_summary = (
            df_app_filtered.groupby("Kode_Rekening")
            .agg({
                "Nama_Belanja": "first",
                "Nilai_Persediaan": "sum"
            })
            .reset_index()
            .rename(
                columns={
                    "Nama_Belanja": "Nama_Rekening_App",
                    "Nilai_Persediaan": "Total_Persediaan"
                }
            )
        )

        recon = pd.merge(lra_summary, app_summary, on="Kode_Rekening", how="outer").fillna(
            {"Total_LRA": 0, "Total_Persediaan": 0}
        )

        # Fallback Nama Rekening
        recon["Nama_Rekening"] = recon["Nama_Rekening_LRA"].combine_first(recon["Nama_Rekening_App"]).fillna("-")
        recon.drop(columns=["Nama_Rekening_LRA", "Nama_Rekening_App"], inplace=True)

        recon["Selisih"] = recon["Total_LRA"] - recon["Total_Persediaan"]

        def hitung_status(row):
            if row["Selisih"] == 0:
                return "✅ Cocok / Balance"
            elif row["Total_Persediaan"] == 0:
                return "⚠️ Belum Diinput di Persediaan"
            elif row["Total_LRA"] == 0:
                return "⚠️ Tidak Ada Realisasi LRA"
            elif row["Selisih"] > 0:
                return "❌ LRA Lebih Besar"
            else:
                return "❌ Persediaan Lebih Besar"

        recon["Status"] = recon.apply(hitung_status, axis=1)

        # Urutkan Kolom
        recon = recon[[
            "Kode_Rekening", "Nama_Rekening", "Total_LRA", "Total_Persediaan", "Selisih", "Status"
        ]].sort_values(by="Kode_Rekening")

        # ================= METRIK & DASHBOARD =================
        st.subheader(f"🏢 {skpd_pilihan}")
        st.caption(f"Periode Data Persediaan: **{periode_label}**")

        tot_lra = recon["Total_LRA"].sum()
        tot_app = recon["Total_Persediaan"].sum()
        tot_selisih = recon["Selisih"].sum()

        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Total Realisasi LRA", format_rupiah(tot_lra))
        col2.metric("Total Aplikasi Persediaan", format_rupiah(tot_app))
        col3.metric(
            "Total Selisih",
            format_rupiah(tot_selisih),
            delta=f"{format_rupiah(tot_selisih)}",
            delta_color="inverse"
        )
        jml_balance = (recon["Status"] == "✅ Cocok / Balance").sum()
        col4.metric("Kesesuaian Akun", f"{jml_balance} / {len(recon)} Akun")

        # ================= TABEL DATA =================
        st.markdown("---")
        st.subheader("📋 Ringkasan Rekonsiliasi Per Kode Rekening")

        # Filter tampilan status
        filter_status = st.multiselect(
            "Filter Status:",
            options=recon["Status"].unique().tolist(),
            default=recon["Status"].unique().tolist()
        )
        recon_view = recon[recon["Status"].isin(filter_status)].copy()

        # Styling tabel dengan warna
        def style_status(row):
            if "Cocok" in row["Status"]:
                return ["background-color: #d4edda; color: #155724;"] * len(row)
            else:
                return ["background-color: #f8d7da; color: #721c24;"] * len(row)

        st.dataframe(
            recon_view.style.apply(style_status, axis=1).format({
                "Total_LRA": format_rupiah,
                "Total_Persediaan": format_rupiah,
                "Selisih": format_rupiah,
            }),
            use_container_width=True,
            height=350
        )

        # ================= DETAIL DRILLDOWN =================
        with st.expander("🔍 Lihat Rincian Transaksi Aplikasi Persediaan (Per Akun)"):
            list_akun = recon_view["Kode_Rekening"].tolist()
            if list_akun:
                akun_terpilih = st.selectbox("Pilih Kode Rekening untuk melihat rincian barang:", list_akun)
                df_detail = df_app_filtered[df_app_filtered["Kode_Rekening"] == akun_terpilih][
                    ["Tanggal", "Referensi", "No_Pesanan", "Kode_Persediaan", "Nama_Persediaan", "Nilai_Persediaan"]
                ]
                st.write(f"Daftar mutasi barang untuk akun **{akun_terpilih}**:")
                st.dataframe(
                    df_detail.style.format({"Nilai_Persediaan": format_rupiah}),
                    use_container_width=True
                )
            else:
                st.info("Tidak ada data akun pada filter yang dipilih.")

        # ================= DOWNLOAD EXCEL =================
        st.markdown("---")
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            recon.to_excel(writer, sheet_name="Ringkasan Rekonsiliasi", index=False)
            df_app_filtered.to_excel(writer, sheet_name="Detail Mutasi Persediaan", index=False)

        st.download_button(
            label="📥 Download Laporan Rekonsiliasi (.xlsx)",
            data=buffer.getvalue(),
            file_name=f"Rekon_{skpd_pilihan.replace(' ', '_')}_{periode_label.replace(' ', '_')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    except Exception as e:
        st.error(f"Terjadi kesalahan saat memproses data: {str(e)}")
        st.exception(e)
else:
    st.info("💡 Silakan unggah ketiga file Excel di menu Sidebar untuk memulai analisis.")

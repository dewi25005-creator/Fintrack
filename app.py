import streamlit as st
import pandas as pd
from datetime import date
from io import BytesIO

# ============================================================
# FUNGSI (FUNCTIONS)
# ============================================================

KOLOM = ["ID", "Tanggal", "Jenis", "Kategori", "Keterangan", "Jumlah"]
KATEGORI = [
    "Gaji", "Usaha", "Beasiswa", "Makanan", "Transportasi",
    "Pendidikan", "Belanja", "Tagihan", "Kesehatan", "Hiburan", "Lainnya"
]
JENIS = ["Pemasukan", "Pengeluaran"]


def format_rupiah(nilai):
    """Mengubah angka menjadi format Rupiah."""
    return "Rp " + f"{nilai:,.0f}".replace(",", ".")


def buat_excel(dataframe):
    """Mengubah DataFrame menjadi file Excel di memori."""
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        dataframe.to_excel(writer, index=False, sheet_name="Transaksi")
    return buffer.getvalue()


def siapkan_data_excel(file):
    """Membaca dan memvalidasi kolom dasar dari file Excel."""
    df = pd.read_excel(file)
    wajib = ["Tanggal", "Jenis", "Kategori", "Keterangan", "Jumlah"]
    if not set(wajib).issubset(df.columns):
        kurang = [kolom for kolom in wajib if kolom not in df.columns]
        raise ValueError("Kolom berikut belum tersedia: " + ", ".join(kurang))

    df = df[wajib].copy()
    df["Tanggal"] = pd.to_datetime(df["Tanggal"], errors="raise").dt.strftime("%Y-%m-%d")
    df["Jenis"] = df["Jenis"].astype(str).str.strip()
    df["Kategori"] = df["Kategori"].astype(str).str.strip()
    df["Keterangan"] = df["Keterangan"].fillna("").astype(str)
    df["Jumlah"] = pd.to_numeric(df["Jumlah"], errors="raise")

    if not df["Jenis"].isin(JENIS).all():
        raise ValueError("Kolom Jenis hanya boleh berisi Pemasukan atau Pengeluaran.")
    if df["Jumlah"].isna().any() or (df["Jumlah"] <= 0).any():
        raise ValueError("Semua Jumlah harus berupa angka positif.")
    return df


# ============================================================
# OOP: CLASS, OBJECT, ATTRIBUTE, DAN METHOD
# ============================================================

class FinanceManager:
    """Mengelola daftar transaksi keuangan."""

    def __init__(self, data=None):
        self.data = data if data is not None else []

    def tambah_transaksi(self, tanggal, jenis, kategori, keterangan, jumlah):
        """Menambahkan satu transaksi baru."""
        nomor_baru = max([item["ID"] for item in self.data], default=0) + 1
        self.data.append({
            "ID": nomor_baru,
            "Tanggal": str(tanggal),
            "Jenis": jenis,
            "Kategori": kategori,
            "Keterangan": keterangan.strip(),
            "Jumlah": float(jumlah),
        })

    def hapus_transaksi(self, id_transaksi):
        """Menghapus transaksi berdasarkan ID."""
        self.data = [
            item for item in self.data
            if int(item["ID"]) != int(id_transaksi)
        ]

    def dataframe(self):
        """Mengubah daftar transaksi menjadi tabel pandas."""
        return pd.DataFrame(self.data, columns=KOLOM)

    def ringkasan(self):
        """Menghitung pemasukan, pengeluaran, dan saldo."""
        pemasukan = sum(
            item["Jumlah"] for item in self.data
            if item["Jenis"] == "Pemasukan"
        )
        pengeluaran = sum(
            item["Jumlah"] for item in self.data
            if item["Jenis"] == "Pengeluaran"
        )
        return pemasukan, pengeluaran, pemasukan - pengeluaran


# ============================================================
# KONFIGURASI DAN TAMPILAN STREAMLIT
# ============================================================

st.set_page_config(
    page_title="FinTrack | Keuangan Pribadi",
    page_icon="💰",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
    .main .block-container {padding-top: 1.8rem; padding-bottom: 2.5rem;}
    [data-testid="stMetric"] {
        background: #f4f7fb;
        border: 1px solid #e4eaf2;
        padding: 16px 18px;
        border-radius: 12px;
    }
    div.stButton > button, div.stDownloadButton > button {
        border-radius: 8px;
        font-weight: 600;
    }
</style>
""", unsafe_allow_html=True)

# Data berada di session_state agar tetap ada selama sesi aplikasi.
if "transaksi" not in st.session_state:
    st.session_state.transaksi = []

# Object dari class FinanceManager.
finance = FinanceManager(st.session_state.transaksi)

st.sidebar.title("💰 FinTrack")
st.sidebar.caption("Pencatatan Keuangan Pribadi")
menu = st.sidebar.radio(
    "Navigasi",
    ["Dashboard", "Tambah Transaksi", "Data Transaksi", "Impor / Ekspor Excel"]
)
st.sidebar.divider()
st.sidebar.info("Catatan: ekspor laporan secara berkala agar data dapat disimpan sebagai file.")

# ------------------------------------------------------------
# MENU 1: DASHBOARD
# ------------------------------------------------------------
if menu == "Dashboard":
    st.title("Dashboard Keuangan")
    st.write("Pantau pemasukan, pengeluaran, dan saldo kamu dalam satu halaman.")

    pemasukan, pengeluaran, saldo = finance.ringkasan()
    c1, c2, c3 = st.columns(3)
    c1.metric("Total Pemasukan", format_rupiah(pemasukan))
    c2.metric("Total Pengeluaran", format_rupiah(pengeluaran))
    c3.metric("Saldo", format_rupiah(saldo))

    df = finance.dataframe()
    st.subheader("Ringkasan Transaksi")

    if df.empty:
        st.info("Belum ada transaksi. Buka menu Tambah Transaksi untuk mulai mencatat.")
    else:
        grafik_col, kategori_col = st.columns(2)

        with grafik_col:
            st.markdown("**Pemasukan vs Pengeluaran**")
            total_jenis = (
                df.groupby("Jenis")["Jumlah"]
                .sum()
                .reindex(["Pemasukan", "Pengeluaran"], fill_value=0)
            )
            st.bar_chart(total_jenis)

        with kategori_col:
            st.markdown("**Pengeluaran berdasarkan kategori**")
            df_keluar = df[df["Jenis"] == "Pengeluaran"]
            if df_keluar.empty:
                st.info("Belum ada transaksi pengeluaran.")
            else:
                total_kategori = df_keluar.groupby("Kategori")["Jumlah"].sum().sort_values(ascending=False)
                st.bar_chart(total_kategori)

        st.subheader("Transaksi Terbaru")
        terbaru = df.copy()
        terbaru["Tanggal"] = pd.to_datetime(terbaru["Tanggal"], errors="coerce")
        terbaru = terbaru.sort_values("Tanggal", ascending=False).head(5)
        terbaru["Tanggal"] = terbaru["Tanggal"].dt.strftime("%Y-%m-%d")
        st.dataframe(terbaru, use_container_width=True, hide_index=True)

# ------------------------------------------------------------
# MENU 2: TAMBAH TRANSAKSI
# ------------------------------------------------------------
elif menu == "Tambah Transaksi":
    st.title("Tambah Transaksi")
    st.write("Isi formulir berikut, lalu tekan tombol simpan.")

    with st.form("form_tambah_transaksi", clear_on_submit=True):
        tanggal = st.date_input("Tanggal transaksi", value=date.today())
        jenis = st.selectbox("Jenis transaksi", JENIS)
        kategori = st.selectbox("Kategori", KATEGORI)
        keterangan = st.text_input("Keterangan", placeholder="Contoh: makan siang")
        jumlah = st.number_input("Jumlah (Rp)", min_value=1, value=10000, step=1000)
        simpan = st.form_submit_button("Simpan Transaksi", type="primary", use_container_width=True)

    if simpan:
        if not keterangan.strip():
            st.error("Keterangan tidak boleh kosong.")
        else:
            finance.tambah_transaksi(tanggal, jenis, kategori, keterangan, jumlah)
            st.session_state.transaksi = finance.data
            st.success("Transaksi berhasil disimpan!")
            st.rerun()

# ------------------------------------------------------------
# MENU 3: DATA TRANSAKSI
# ------------------------------------------------------------
elif menu == "Data Transaksi":
    st.title("Data Transaksi")
    df = finance.dataframe()

    if df.empty:
        st.info("Belum ada data transaksi.")
    else:
        filter_col, search_col = st.columns([1, 2])
        with filter_col:
            filter_jenis = st.selectbox("Filter jenis", ["Semua"] + JENIS)
        with search_col:
            kata_kunci = st.text_input("Cari keterangan atau kategori", placeholder="Ketik kata kunci...")

        hasil = df.copy()
        if filter_jenis != "Semua":
            hasil = hasil[hasil["Jenis"] == filter_jenis]
        if kata_kunci.strip():
            mask = (
                hasil["Keterangan"].str.contains(kata_kunci, case=False, na=False)
                | hasil["Kategori"].str.contains(kata_kunci, case=False, na=False)
            )
            hasil = hasil[mask]

        st.caption(f"Menampilkan {len(hasil)} dari {len(df)} transaksi.")
        st.dataframe(hasil, use_container_width=True, hide_index=True)

        st.subheader("Hapus Transaksi")
        pilihan_id = st.selectbox(
            "Pilih ID transaksi yang akan dihapus",
            options=df["ID"].tolist(),
            format_func=lambda x: (
                f"ID {x} — {df.loc[df['ID'] == x, 'Tanggal'].iloc[0]} — "
                f"{df.loc[df['ID'] == x, 'Keterangan'].iloc[0]}"
            ),
        )
        konfirmasi = st.checkbox("Saya yakin ingin menghapus transaksi ini.")
        if st.button("Hapus Transaksi", type="primary"):
            if konfirmasi:
                finance.hapus_transaksi(pilihan_id)
                st.session_state.transaksi = finance.data
                st.success("Transaksi berhasil dihapus.")
                st.rerun()
            else:
                st.warning("Centang konfirmasi terlebih dahulu.")

# ------------------------------------------------------------
# MENU 4: IMPOR / EKSPOR EXCEL
# ------------------------------------------------------------
elif menu == "Impor / Ekspor Excel":
    st.title("Impor dan Ekspor Excel")
    st.write("Gunakan file Excel dengan kolom Tanggal, Jenis, Kategori, Keterangan, dan Jumlah.")

    st.subheader("Impor data")
    file_excel = st.file_uploader("Pilih file Excel (.xlsx)", type=["xlsx"])

    if file_excel is not None:
        try:
            df_impor = siapkan_data_excel(file_excel)
            st.write("Pratinjau data yang akan diimpor:")
            st.dataframe(df_impor, use_container_width=True, hide_index=True)

            if st.button("Impor Transaksi", type="primary"):
                # Beri ID baru supaya tidak bentrok dengan transaksi yang sudah ada.
                for _, baris in df_impor.iterrows():
                    finance.tambah_transaksi(
                        baris["Tanggal"],
                        baris["Jenis"],
                        baris["Kategori"],
                        baris["Keterangan"],
                        baris["Jumlah"],
                    )
                st.session_state.transaksi = finance.data
                st.success(f"{len(df_impor)} transaksi berhasil diimpor.")
                st.rerun()
        except Exception as error:
            st.error(f"File tidak dapat diproses: {error}")

    st.divider()
    st.subheader("Ekspor laporan")
    df_ekspor = finance.dataframe()
    st.write(f"Jumlah transaksi saat ini: **{len(df_ekspor)}**")
    st.download_button(
        label="⬇️ Unduh Laporan Excel",
        data=buat_excel(df_ekspor),
        file_name="laporan_keuangan_fintrack.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

st.divider()
st.caption("FinTrack • Proyek UTS Python • Data tersimpan selama sesi aplikasi berlangsung.")

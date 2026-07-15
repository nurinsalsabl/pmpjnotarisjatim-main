import streamlit as st
import pandas as pd
import json
from datetime import datetime
import os
import openpyxl
import pytesseract
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from io import BytesIO
from PyPDF2 import PdfReader
import difflib
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from google_auth_oauthlib.flow import InstalledAppFlow
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
import pdfplumber
import io
from PIL import Image


st.set_page_config(
    page_title="Kuisioner PMPJ Notaris - Kanwil Kemenkumham Jawa Timur",
    page_icon="⚖️",
    layout="wide",
    initial_sidebar_state="expanded",
)

from ui_styles import load_css, render_header, render_sidebar_stepper, render_progress_bar, render_report_table

load_css()

scope = [
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/drive.file",
    "https://www.googleapis.com/auth/spreadsheets"
]

# =========================================================
# AUTENTIKASI GOOGLE (hybrid: lokal + Streamlit Cloud)
# DIPERBAIKI: dibungkus try/except di setiap tahap supaya
# satu kegagalan (token rusak, secrets belum ada, dsb) tidak
# menjatuhkan seluruh aplikasi sebelum wizard sempat tampil.
# =========================================================
creds = None

if os.path.exists("token.json"):
    try:
        creds = Credentials.from_authorized_user_file("token.json", scopes=scope)
    except Exception:
        creds = None

if creds is None:
    try:
        if "google" in st.secrets and "token" in st.secrets["google"]:
            creds_data = json.loads(st.secrets["google"]["token"])
            creds = Credentials.from_authorized_user_info(creds_data, scopes=scope)
    except Exception:
        # Tidak ada secrets.toml sama sekali (mis. dijalankan lokal) -> abaikan, jangan crash
        creds = None

if creds and creds.expired and creds.refresh_token:
    try:
        creds.refresh(Request())
    except Exception:
        pass

if not creds:
    st.warning("🔐 Token belum ada, buka login Google OAuth untuk membuat token.json (hanya di lokal).")
    try:
        flow = InstalledAppFlow.from_client_secrets_file("credentials.json", scopes=scope)
        creds = flow.run_local_server(port=0)
        with open("token.json", "w") as token_file:
            token_file.write(creds.to_json())
    except Exception as e:
        #st.error(f"❌ Tidak dapat melakukan otentikasi Google: {e}")
        creds = None

try:
    client = gspread.authorize(creds) if creds else None
except Exception:
    client = None

try:
    drive_service = build('drive', 'v3', credentials=creds) if creds else None
except Exception:
    drive_service = None

if client:
    try:
        sh = client.open("Kuisioner PMPJ Notaris FINAL 2026")
        worksheet = sh.sheet1
        st.success("📄 Mohon lengkapi kuisioner berikut sesuai format!")
    except gspread.SpreadsheetNotFound:
        st.error("❌ Spreadsheet tidak ditemukan. Pastikan sudah dibagikan ke akun Google.")
        sh = None
        worksheet = None
    except Exception as e:
        # DIPERBAIKI: sebelumnya hanya menangkap gspread.SpreadsheetNotFound,
        # sehingga error jaringan/otentikasi lain membuat seluruh app crash.
        #st.error(f"❌ Gagal membuka spreadsheet: {e}")
        sh = None
        worksheet = None
else:
    sh = None
    worksheet = None

# =========================================================
# DICTIONARY MAPPING (TIDAK DIUBAH)
# =========================================================
profil = {
    "a. Pengusaha/wiraswasta": 9,
    "b.  PNS (termasuk pensiunan)": 4,
    "c.  Ibu Rumah Tangga": 2,
    "d.  Pelajar/Mahasiswa": 2,
    "e.  Pegawai Swasta": 7,
    "f.  Pejabat Lembaga Legislatif dan Pemerintah": 4,
    "g.  TNI/POLRI (termasuk Pensiunan)": 3,
    "h. Pegawai BI/BUMN/BUMD (termasuk Pensiunan)": 2,
    "i.  Profesional dan Konsultan": 6,
    "j.  Pedagang": 5,
    "k.  Pegawai Bank": 2,
    "l. Pegawai Money Changer": 1,
    "m. Pengajar dan Dosen": 2,
    "n. Petani": 1,
    "o.  Korporasi Perseroan Terbatas": 7,
    "p.  Korporasi Koperasi": 2,
    "q.  Korporasi Yayasan": 2,
    "r.  Korporasi CV, Firma, dan Maatschap": 2,
    "s.  Korporasi Perkumpulan Badan Hukum": 2,
    "t.  Korporasi Perkumpulan Tidak Badan Hukum": 2,
    "u.  Pengurus Parpol": 2,
    "v.  Bertindak berdasarkan Kuasa": 2,
    "w. Lain-lain": 1
}

bisnis_pengguna = {
    "a. Perdagangan": 9,
    "b. Pertambangan": 4,
    "c. Pertanian": 1,
    "d. Perikanan": 1,
    "e. Perkebunan": 1,
    "f. Perindustrian": 2,
    "g. Perbankan": 3,
    "h. Pembiayaan": 4,
    "i. Pembangunan Property": 3,
    "j. Kontraktor": 2,
    "k. Konsultan": 1,
    "l. Transportasi Barang dan Orang": 1,
    "m. Usaha Sewa Menyewa": 2,
    "n. Lain-lain....": 1
}

jasa = {
    "a.  Pembelian dan Penjualan Properti": 9,
    "b.  Pengurusan Perizinan Badan Usaha": 7,
    "c.  Penitipan Pembayaran Pajak terkait Pengalihan Property": 3,
    "d.  Pengurusan Pembelian dan Penjualan Badan Usaha": 3,
    "e.  Pengelolaan terhadap Uang, Efek, dan/atau Produk Jasa Keuangan lainnya": 4,
    "f.  Pengelolaan Rekening Giro, Rekening Tabungan, Rekening Deposito, dan/atau Rekening Efek": 2,
    "g.  Pengoperasian dan Pengelolaan Perusahaan": 3,
    "h. Lain-lain": 1
}

produk = {
    "a. Akta pembayaran uang sewa, bunga, dan pensiun ": 4,
    "b. Akta penawaran pembayaran tunai ": 4,
    "c.  Akta protes terhadap tidak dibayarnya atau tidak diterimanya surat berharga ": 2,
    "d. Akta Kuasa": 4,
    "e. Akta keterangan kepemilikan": 5,
    "f. Akta Hibah (Barang Bergerak)": 4,
    "g. Akta Wasiat": 2,
    "h. Akta Jaminan Fidusia ": 3,
    "i. Akta Pendirian Perseroan Terbatas ": 8,
    "j. Akta Perubahan Perseroan Terbatas  ": 5,
    "k. Akta Pendirian dan Perubahan Koperasi ": 3,
    "l. Akta Pendirian dan Perubahan Yayasan (Nirlaba) ": 3,
    "m. Akta Pendirian dan Perubahan CV, Firma dan Maatschap (Persekutuan Perdata) - Badan usaha yang tidak berbadan hukum ": 3,
    "n. Akta Pendirian dan Perubahan Perkumpulan Badan Hukum (Sosial/Nirlaba) ": 3,
    "o. Akta Pendirian dan Perubahan Perkumpulan Tidak Berbadan Hukum (Sosial/Nirlaba) ": 3,
    "p. Akta Pendirian dan Perubahan Partai Politik ": 2,
    "q. Akta Perjanjian Sewa Menyewa ": 3,
    "r. Akta Perjanjian Pengikatan Jual Beli ": 8,
    "s. Akta Perjanjian Kerjasama ": 4,
    "t. Akta Perjanjian BOT (Build Operate Transfer/Bangun Kelola Serah) ": 2,
    "u. Akta Perjanjian JO (Joint Operation/Kerjasama Operasional Mengelola Proyek) ": 2,
    "v. Akta Perjanjian Kredit ": 4,
    "w. Akta Pinjam Meminjam/Pengakuan Hutang ": 4,
    "x. Akta lainnya sesuai dengan ketentuan peraturan perundang-undangan ": 3
}

negara = {
    "a.  Tax Haven Country": 6,
    "b.  RRT (Tiongkok)": 8,
    "c.  Malaysia": 7,
    "d.  Singapura": 7,
    "e.  Asia lainnya": 8,
    "f.  Afrika": 1,
    "g.  Amerika": 5,
    "h.  Eropa": 6,
    "i.  Australia dan Selandia Baru": 5
}

apgakkum = {"YA": 6, "TIDAK": 1}

wilayah_skor = {
    "Jawa Timur": 6, "Jawa Tengah": 4, "Jawa Barat": 6, "DKI Jakarta": 9, "Aceh": 5,
    "Kalimantan Timur": 4, "Banten": 3, "Kepulauan Riau": 3, "Lampung": 3,
    "Sulawasi Selatan": 3, "Sumatera Utara": 3, "Sulawasi Tenggara": 3, "Sulawesi Utara": 3,
    "Sumatera Selatan": 3, "DI Yogyakarta": 3, "Bali": 2, "Riau": 2, "Bangka Belitung": 2,
    "Bengkulu": 2, "Kalimantan Tengah": 2, "Maluku Utara": 2, "Nusa Tenggara Timur": 2,
    "Papua": 2, "Sulawesi Barat": 2, "Sulawesi Tengah": 2, "Gorontalo": 2, "Jambi": 2,
    "Kalimantan Selatan": 2, "Maluku": 2, "Nusa Tenggara Barat": 2, "Papua Barat": 2,
    "Sumatera Barat": 2, "Kalimantan Barat": 1, "Kalimantan Utara": 1
}

daftar_kota = [
    "Kabupaten Bangkalan", "Kabupaten Banyuwangi", "Kabupaten Blitar", "Kabupaten Bojonegoro",
    "Kabupaten Bondowoso", "Kabupaten Gresik", "Kabupaten Jember", "Kabupaten Jombang",
    "Kabupaten Kediri", "Kabupaten Lamongan", "Kabupaten Lumajang", "Kabupaten Madiun",
    "Kabupaten Magetan", "Kabupaten Malang", "Kabupaten Mojokerto", "Kabupaten Nganjuk",
    "Kabupaten Ngawi", "Kabupaten Pacitan", "Kabupaten Pamekasan", "Kabupaten Pasuruan",
    "Kabupaten Ponorogo", "Kabupaten Probolinggo", "Kabupaten Sampang", "Kabupaten Sidoarjo",
    "Kabupaten Situbondo", "Kabupaten Sumenep", "Kabupaten Trenggalek", "Kabupaten Tuban",
    "Kabupaten Tulungagung", "Kota Batu", "Kota Blitar", "Kota Kediri", "Kota Madiun",
    "Kota Malang", "Kota Mojokerto", "Kota Pasuruan", "Kota Probolinggo", "Kota Surabaya"
]

daftar_wilayah = list(wilayah_skor.keys())

q_cols = [
    "1.  Apakah Kantor Notaris anda memiliki mekanisme analisis risiko Pengguna Jasa? (form cdd, edd dan analisa resiko)",
    "2.  Apakah Kantor Notaris anda memiliki kebijakan dan prosedur untuk mengelola dan memitigasi risiko tinggi ... Pasal 17 PerMenkumham 9/2017?",
    "3.  Apakah Kantor Notaris anda memiliki kebijakan dan prosedur untuk mengelola dan memitigasi risiko sedang pencucian uang dan/atau pendanaan terorisme sebagaimana diatur dalam Pasal 7 PerMenkumham Nomor 9 Tahun 2017?",
    "4.  Apakah Kantor Notaris anda memiliki kebijakan dan prosedur untuk mengelola dan memitigasi risiko rendah pencucian uang dan/atau pendanaan terorisme sebagaimana diatur dalam Pasal 16 PerMenkumham Nomor 9 Tahun 2017?",
    "5.  Kebijakan larangan nama fiktif?",
    "6.  Pengumpulan informasi OP (Pasal 7 ayat 1)?",
    "7.  Pengumpulan informasi Korporasi (Pasal 7 ayat 2)?",
    "8.  Pengumpulan informasi legal arrangements?",
    "9.  Pengumpulan BO Korporasi (Pasal 8)?",
    "10. Pengumpulan BO legal arrangement (Pasal 9)?",
    "11. Kebijakan bertemu langsung dengan pengguna jasa?",
    "12. Konfirmasi kebenaran formil dokumen ke pengguna jasa?",
    "13. Meminta dokumen pendukung dari otoritas berwenang bila ragu?",
    "14. Pemantauan kewajaran transaksi?",
    "15. Pencatatan transaksi pengguna jasa?",
    "16. Sistem informasi identifikasi/pemantauan/laporan transaksi?",
    "17. Penatausahaan dokumen transaksi & BO?",
    "18. Penatausahaan dokumen PJP & BO (prinsip PJP)?",
    "19. Penatausahaan dokumen analisis kewajaran transaksi?",
    "20. Putus hubungan bila PJP menolak prinsip PJP?",
    "21. Putus hubungan bila info PJP diragukan?",
    "22. Laporkan ke PPATK atas tindakan di no.20 & no.21?",
    "23. Pemutakhiran informasi & dokumen PJP bila ada perubahan?",
    "24. Dokumentasi hasil pemutakhiran?",
    "25. Pemeriksaan berkala penerapan prinsip PJP?",
    "26. Pemutakhiran daftar PJP/kuasa berisiko tinggi?",
    "27. Screening penerimaan karyawan (pre-employee)?",
    "28. Pengenalan & pemantauan profil karyawan?",
    "29. Sosialisasi/perlatihan regulasi PJP?",
    "30. Sosialisasi/perlatihan tipologi TPPU/TPPT?",
    "31. Sosialisasi/perlatihan kebijakan & prosedur PJP?",
    "32. Pemanfaatan teknologi baru?",
    "33. Identifikasi & pengukuran risiko sebelum adopsi teknologi baru?",
    "34. Kerja sama dengan penegak hukum & otoritas berwenang?"
]
# Teks pertanyaan lengkap (dipakai untuk ditampilkan di form kepatuhan)
q_full_text = {
    1: "1. Apakah Kantor Notaris anda memiliki mekanisme analisis risiko Pengguna Jasa? (form cdd, edd dan analisa resiko)?",
    2: "2.  Apakah Kantor Notaris anda memiliki kebijakan dan prosedur untuk mengelola dan memitigasi risiko tinggi pencucian uang dan/atau pendanaan terorisme, termasuk PEP dan negara yang berisiko tinggi sebagaimana diatur dalam Pasal 17 PerMenkumham Nomor 9 Tahun 2017?",
    3: "3.  Apakah Kantor Notaris anda memiliki kebijakan dan prosedur untuk mengelola dan memitigasi risiko sedang pencucian uang dan/atau pendanaan terorisme sebagaimana diatur dalam Pasal 7 PerMenkumham Nomor 9 Tahun 2017?",
    4: "4.  Apakah Kantor Notaris anda memiliki kebijakan dan prosedur untuk mengelola dan memitigasi risiko rendah pencucian uang dan/atau pendanaan terorisme sebagaimana diatur dalam Pasal 16 PerMenkumham Nomor 9 Tahun 2017?",
    5: "5.  Apakah Kantor Notaris Anda memiliki kebijakan larangan untuk membuka atau memelihara hubungan usaha yang menggunakan nama fiktif?",
    6: "6.  Apakah Kantor Notaris Anda melakukan pengumpulan informasi pengguna jasa orang perseorangan sebagaimana dimaksud dalam Pasal 7 ayat (1) PerMenkumham Nomor 9 Tahun 2017?",
    7: "7.  Apakah Kantor Notaris Anda melakukan pengumpulan informasi pengguna jasa korporasi sebagaimana dimaksud dalam Pasal 7 ayat (2) PerMenkumham Nomor 9 Tahun 2017?",
    8: "8.  Apakah Kantor Notaris Anda melakukan pengumpulan informasi pengguna jasa perikatan lainnya (legal arrangements) sebagaimana dimaksud dalam Pasal 7 ayat (1) PerMenkumham Nomor 9 Tahun 2017?",
    9: "9.  Apakah Kantor Notaris Anda melakukan pengumpulan informasi Beneficial Owner (Pemilik Manfaat) dari Korporasi sebagaimana dimaksud dalam Pasal 8 PerMenkumham Nomor 9 Tahun 2017?",
    10: "10.  Apakah Kantor Notaris Anda melakukan pengumpulan informasi Beneficial Owner (Pemilik Manfaat) dari perikatan lainnya (legal arrangement) sebagaimana dimaksud dalam Pasal 9 PerMenkumham Nomor 9 Tahun 2017?",
    11: "11.  Apakah Kantor Notaris Anda memliliki kebijakan bertemu langsung dengan pegguna jasa dalam rangka pengumpulan informasi pengguna jasa?",
    12: "12. Apakah Kantor Notaris Anda melakukan konfirmasi atas dokumen Pengguna Jasa melalui meminta keterangan kepada pengguna jasa untuk mengetahui kebenaran formil dokumen dimaksud?",
    13: "13. Apakah Kantor Notaris Anda memiliki kebijakan untuk meminta dokumen pendukung lainnya dari pihak yang berwenang dalam hal terdapat keraguan atas kebenaran formil dokumen?",
    14: "14. Apakah Kantor Notaris anda melakukan pemantauan kewajaran transaksi pengguna jasa?",
    15: "15. Apakah Kantor Notaris anda melakukan pencatatan transaksi pengguna jasa?",
    16: "16. Apakah Kantor Notaris Anda memiliki sistem informasi mengenai identifikasi, pemantauan transaksi, dan penyediaan laporan mengenai transaksi yang dilakukan oleh pengguna jasa?",
    17: "17. Apakah Kantor Notaris anda menatausahakan dokumen seluruh informasi dan dokumen transaksi pengguna jasa dan beneficial owner (pemilik manfaat)?",
    18: "18. Apakah Kantor Notaris anda menatausahakan dokumen seluruh informasi dan dokumen pengguna jasa dan beneficial owner (pemilik manfaat) yang diperoleh dalam rangka penerapan prinsip mengenali pengguna jasa?",
    19: "19. Apakah Kantor Notaris anda menatausahakan dokumen analisis kewajaran atas transaksi pengguna jasa dan beneficial owner (pemilik manfaat)?",
    20: "20. Apakah Kantor Notaris anda akan memutuskan hubungan usaha dengan pengguna jasa jika Pengguna Jasa menolak untuk mematuhi prinsip mengenali Pengguna Jasa?",
    21: "21. Apakah Kantor Notaris anda akan memutuskan hubungan usaha dengan pengguna jasa jika Notaris meragukan kebenaran informasi yang disampaikan oleh Pengguna Jasa?",
    22: "22. Apakah Kantor Notaris anda akan melaporkan kepada PPATK mengenai tindakan sebagaimana dimaksud pada pertanyaan nomor 20 dan nomor 21 di atas?",
    23: "23. Apakah Kantor Notaris anda melakukan upaya pemutakhiran informasi dan dokumen pengguna jasa dalam hal terdapat perubahan yang diketahui oleh Notaris yang bersumber dari Pengguna Jasa yang sama atau informasi lain yang dapat dipertanggungjawabkan?",
    24: "24. Apakah Kantor Notaris Anda mendokumentasikan hasil pemutakhiran informasi dan/atau dokumen sebagaimana dimaksud dalam pertanyaan Nomor 23?",
    25: "25. Apakah Kantor Notaris Anda melakukan pengendalian internal melalui pelaksanaan pemeriksaan berkala terhadap penerapan prinsip mengenali Pengguna Jasa?",
    26: "26. Apakah Kantor Notaris Anda melakukan pengendalian internal melalui Pemutakhiran daftar Pengguna Jasa atau pemberi kuasa yang memenuhi kriteria berisiko tinggi?",
    27: "27. Apakah Kantor Notaris Anda melakukan prosedur penyaringan untuk penerimaan karyawan baru (pre-employee screening)?",
    28: "28. Apakah Kantor Notaris Anda melakukan pengenalan dan pemantauan terhadap profil karyawan?",
    29: "29. Apakah Kantor Notaris anda melakukan sosialisasi atau pelatihan mengenai penerapan peraturan perundang-undangan yang terkait dengan prinsip mengenali Pengguna Jasa? (baik diselenggarakan secara mandiri atau instansi terkait)",
    30: "30. Apakah Kantor Notaris anda melakukan sosialisasi atau pelatihan mengenai teknik, metode, dan tipologi pencucian uang dan/atau pendanaan terorisme? (baik diselenggarakan secara mandiri atau instansi terkait)",
    31: "31. Apakah Kantor Notaris anda melakukan sosialisasi atau pelatihan mengenai kebijakan dan prosedur penerapan prinsip mengenali Pengguna Jasa serta peran dan tanggung jawab pegawai dalam mencegah dan memberantas pencucian uang dan/atau pendanaan terorisme? (baik diselenggarakan secara mandiri atau instansi terkait)",
    32: "32. Apakah Kantor Notaris Anda memanfaatkan teknologi baru dalam memberikan pelayanan kepada pengguna jasa?",
    33: "33. Apabila jawaban Nomor 32 adalah iya, apakah Kantor Notaris Anda melakukan pengidentifikasian dan pengukuran mengenai risiko terjadinya tindak pidana pencucian uang dan pendanaan terorisme, sebelum pemanfaatan atau pengembangan teknologi baru tersebut pada pertanyaan Nomor 32?",
    34: "34.  Apakah Kantor Notaris anda pernah melakukan kerja sama dengan penegak hukum dan otoritas yang berwenang untuk memberantas tindak pidana pencucian uang dan pendanaan terorisme?"
}

# =========================================================
# FUNGSI PERHITUNGAN RISIKO (TIDAK DIUBAH)
# =========================================================
def hitung_risiko(inputs):
    def pilih_terbesar(mapping_dict, user_inputs, default=None):
        if all(v == 0 for v in user_inputs.values()):
            return default, mapping_dict.get(default, 0)
        terbaik = max(user_inputs, key=user_inputs.get)
        return terbaik, mapping_dict.get(terbaik, 0)

    jawaban_profil, skor_profil = pilih_terbesar(profil, inputs["profil"], default="w. Lain-lain")
    jawaban_bisnis, skor_bisnis = pilih_terbesar(bisnis_pengguna, inputs["bisnis"], default="n. Lain-lain....")
    jawaban_jasa, skor_jasa = pilih_terbesar(jasa, inputs["jasa"], default="h. Lain-lain")
    jawaban_negara, skor_negara = pilih_terbesar(negara, inputs["negara"], default="e.  Asia lainnya")
    skor_apgakkum = apgakkum.get(inputs["apgakkum"], 0)
    jawaban_wilayah = inputs["wilayah"]
    skor_wilayah = wilayah_skor.get(jawaban_wilayah, 0)

    total = skor_profil + skor_bisnis + skor_jasa + skor_negara + skor_apgakkum + skor_wilayah

    if 6 <= total <= 17: kategori = "Rendah"
    elif 18 <= total <= 29: kategori = "Sedang"
    elif 30 <= total <= 41: kategori = "Tinggi"
    elif 42 <= total <= 52: kategori = "Sangat Tinggi"
    else: kategori = "Diluar Rentang"

    return {
        "jawaban_profil": jawaban_profil, "skor_profil": skor_profil,
        "jawaban_bisnis": jawaban_bisnis, "skor_bisnis": skor_bisnis,
        "jawaban_jasa": jawaban_jasa, "skor_jasa": skor_jasa,
        "jawaban_negara": jawaban_negara, "skor_negara": skor_negara,
        "jawaban_apgakkum": inputs["apgakkum"], "skor_apgakkum": skor_apgakkum,
        "jawaban_wilayah": jawaban_wilayah, "skor_wilayah": skor_wilayah,
        "total_skor": total, "kategori_risiko": kategori
    }

def validasi_ocr_pdf(uploaded_file1, kata_kunci_list, judul=""):
    if uploaded_file1 is None:
        return False, "Tidak ada file.", 0
    try:
        pdf_bytes = uploaded_file1.read()
        uploaded_file1.seek(0)
        all_text = ""
        try:
            with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
                for page in pdf.pages[:5]:
                    extracted = page.extract_text()
                    if extracted:
                        all_text += extracted.lower() + "\n"
        except Exception as e:
            st.warning(f"⚠️ Gagal ekstrak teks langsung: {e}")

        if not all_text.strip():
            st.info("📸 Proses scan PDF sedang berjalan...")
            try:
                with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
                    for i, page in enumerate(pdf.pages[:5]):
                        page_image = page.to_image(resolution=200).original
                        text = pytesseract.image_to_string(page_image, lang="ind+eng", config="--psm 6")
                        all_text += text.lower() + "\n"
            except Exception as e:
                st.error(f"❌ Gagal OCR dari gambar PDF: {e}")
                return False, "Error OCR", 0

        variasi_kata = [
            "formulir customer due diligence perorangan",
            "formulir customer due diligence",
            "enhanced due diligence", "formulir customer due diligence korporasi"
        ]

        jumlah_ditemukan = 0
        for kata_utama in kata_kunci_list:
            kata_lower = kata_utama.lower()
            variasi_relevan = [v for v in variasi_kata if kata_lower in v.lower()]

            def fuzzy_found(keyword):
                panjang = len(keyword)
                for i in range(0, len(all_text) - panjang + 1):
                    potongan = all_text[i:i + panjang + 3]
                    if difflib.SequenceMatcher(None, keyword, potongan).ratio() > 0.6:
                        return True
                return False

            found = (
                kata_lower in all_text
                or any(v in all_text for v in variasi_relevan)
                or fuzzy_found(kata_lower)
            )
            if found:
                jumlah_ditemukan += 1

        if not all_text.strip():
            st.warning(f"⚠️ OCR tidak menemukan teks di file {judul}.")
            return False, "Tidak ada teks terdeteksi", 0

        return True, all_text, jumlah_ditemukan
    except Exception as e:
        st.error(f"❌ Error umum saat OCR: {e}")
        return False, "Error OCR", 0

# =========================================================
# NILAI INTERNAL CONTROL (TIDAK DIUBAH)
# =========================================================

IC_SCORE = {
    1: {"YA": 1, "TIDAK": 5},
    2: {"YA": 1, "TIDAK": 5},
    3: {"YA": 2, "TIDAK": 5},
    4: {"YA": 2, "TIDAK": 5},
    5: {"YA": 1, "TIDAK": 5},
    6: {"YA": 1, "TIDAK": 6},
    7: {"YA": 1, "TIDAK": 6},
    8: {"YA": 1, "TIDAK": 5},
    9: {"YA": 1, "TIDAK": 6},
    10: {"YA": 1, "TIDAK": 5},
    11: {"YA": 1, "TIDAK": 6},
    12: {"YA": 1, "TIDAK": 2},
    13: {"YA": 1, "TIDAK": 6},
    14: {"YA": 1, "TIDAK": 6},
    15: {"YA": 1, "TIDAK": 5},
    16: {"YA": 1, "TIDAK": 5},
    17: {"YA": 1, "TIDAK": 5},
    18: {"YA": 1, "TIDAK": 5},
    19: {"YA": 1, "TIDAK": 5},
    20: {"YA": 1, "TIDAK": 6},
    21: {"YA": 1, "TIDAK": 6},
    22: {"YA": 1, "TIDAK": 5},
    23: {"YA": 2, "TIDAK": 2},
    24: {"YA": 2, "TIDAK": 2},
    25: {"YA": 1, "TIDAK": 5},
    26: {"YA": 2, "TIDAK": 2},
    27: {"YA": 2, "TIDAK": 2},
    28: {"YA": 2, "TIDAK": 2},
    29: {"YA": 1, "TIDAK": 5},
    30: {"YA": 2, "TIDAK": 2},
    31: {"YA": 2, "TIDAK": 2},
    32: {"YA": 1, "TIDAK": 1},
    33: {"YA": 1, "TIDAK": 5},
}

# --- Fungsi hitung internal control ---
def hitung_internal_control(q_answers, uploaded_file1, is_valid_ocr_q1):

    total = 0
    # ==========================
    # Q1 (pakai OCR)
    # ==========================
    if q_answers[1] == "YA":
        if uploaded_file1 is not None and is_valid_ocr_q1:
            total += 1
        else:
            total += 5
    else:
        total += 5
    # ==========================
    # Q2 - Q33
    # ==========================
    for i in range(2, 34):

        jawaban = q_answers.get(i, "TIDAK")

        total += IC_SCORE[i][jawaban]
    # ==========================
    # Kategori
    # ==========================
    if 37 <= total <= 62:
        kategori = "Sangat Baik"
    elif 63 <= total <= 88:
        kategori = "Baik"
    elif 89 <= total <= 114:
        kategori = "Cukup"
    elif 115 <= total <= 141:
        kategori = "Lemah"
    else:
        kategori = "Diluar Rentang"
    return total, kategori

# --- Residual Risk ---
def hitung_residual_risk(kategori_inherent, kategori_internal):
    residual_matrix = {
        "Lemah": {"Rendah": "Rendah", "Sedang": "Sedang", "Tinggi": "Sangat Tinggi", "Sangat Tinggi": "Sangat Tinggi"},
        "Cukup": {"Rendah": "Rendah", "Sedang": "Sedang", "Tinggi": "Tinggi", "Sangat Tinggi": "Sangat Tinggi"},
        "Baik": {"Rendah": "Rendah", "Sedang": "Sedang", "Tinggi": "Sedang", "Sangat Tinggi": "Tinggi"},
        "Sangat Baik": {"Rendah": "Rendah", "Sedang": "Rendah", "Tinggi": "Sedang", "Sangat Tinggi": "Tinggi"}
    }
    risk_value = {"Rendah": 1, "Sedang": 2, "Tinggi": 3, "Sangat Tinggi": 4}
    kategori_residual = residual_matrix.get(kategori_internal, {}).get(kategori_inherent, "Sangat Tinggi")
    return kategori_residual, risk_value.get(kategori_residual, 4)

# --- Risiko Pengguna Jasa ---
def risiko_pengguna_jasa(jumlah_klien):
    if jumlah_klien <= 100: return 1, "Rendah"
    if jumlah_klien <= 200: return 2, "Sedang"
    if jumlah_klien <= 300: return 3, "Tinggi"
    return 4, "Sangat Tinggi"

# --- Final Risk Priority ---
def final_risk(df):
    risk_priority = {
        4: {1: "Tinggi", 2: "Tinggi", 3: "Sangat Tinggi", 4: "Sangat Tinggi"},
        3: {1: "Sedang", 2: "Sedang", 3: "Tinggi", 4: "Sangat Tinggi"},
        2: {1: "Rendah", 2: "Sedang", 3: "Sedang", 4: "Tinggi"},
        1: {1: "Rendah", 2: "Rendah", 3: "Sedang", 4: "Tinggi"}
    }
    df["Tingkat Risiko"] = df.apply(lambda r: risk_priority.get(r["Nilai Residual Risk"], {}).get(r["Nilai Risiko Pengguna Jasa"]), axis=1)
    return df

def colnum_to_excel(n: int) -> str:
    result = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        result = chr(65 + remainder) + result
    return result

# =========================================================
# STATE WIZARD
# =========================================================
STEP_TITLES = [
    "Identitas Notaris",
    "Klien sesuai Profesi",
    "Klien sesuai Bisnis",
    "Jasa & Produk yang Diurus",
    "Klien sesuai Negara & APGAKKUM",
    "Kepatuhan (No. 1-10)",
    "Kepatuhan (No. 11-22)",
    "Kepatuhan (No. 23-34)",
    "Ringkasan & Kirim",
]
TOTAL_STEPS = len(STEP_TITLES)

if "step" not in st.session_state:
    st.session_state.step = 1

def go_next():
    st.session_state.step = min(st.session_state.step + 1, TOTAL_STEPS)

def go_prev():
    st.session_state.step = max(st.session_state.step - 1, 1)

def go_to(n):
    st.session_state.step = n

# =========================================================
# PENYIMPANAN JAWABAN PERMANEN (PERBAIKAN UTAMA)
# =========================================================
# Streamlit menghapus session_state milik sebuah widget begitu widget itu
# TIDAK ikut dirender pada suatu run (mis. karena kita sudah pindah step).
# Wizard bertahap ini hanya merender widget step yang sedang aktif, sehingga
# pola lama (mengandalkan st.session_state["key_widget"] langsung di step 9)
# membuat semua isian di step-step sebelumnya HILANG begitu berpindah step.
#
# Perbaikan: setiap jawaban disalin ke dictionary permanen `st.session_state.answers`
# TEPAT SETELAH widget-nya dirender, pada setiap step. Dictionary permanen ini
# BUKAN milik satu widget manapun sehingga tidak pernah dihapus otomatis oleh
# Streamlit, dan itulah yang dipakai untuk ringkasan maupun pengiriman ke Sheets.
if "answers" not in st.session_state:
    st.session_state.answers = {
        "identitas": {
            "nama_notaris": "", "NIK_KTP": "", "username": "", "nomor_HP": "",
            "alamat": "", "kota": daftar_kota[0], "wilayah_input": daftar_wilayah[0],
        },
        "profil": {k: 0 for k in profil.keys()},
        "bisnis": {k: 0 for k in bisnis_pengguna.keys()},
        "jasa": {k: 0 for k in jasa.keys()},
        "produk": {k: 0 for k in produk.keys()},
        "negara": {k: 0 for k in negara.keys()},
        "apgakkum": "TIDAK",
        "q": {i: "YA" for i in range(1, 35)},
        "file1": None,
        "file2": None,
    }
answers = st.session_state.answers

render_header()
render_sidebar_stepper(STEP_TITLES, st.session_state.step)
render_progress_bar(STEP_TITLES, st.session_state.step)

# =========================================================
# STEP 1 — IDENTITAS NOTARIS
# =========================================================
if st.session_state.step == 1:
    st.markdown('<div class="step-card">', unsafe_allow_html=True)
    st.markdown("<h2>1. Identitas Notaris</h2>", unsafe_allow_html=True)
    st.markdown('<p class="step-subtext">Lengkapi identitas Kantor Notaris Anda. Seluruh kolom wajib diisi.</p>', unsafe_allow_html=True)

    ident = answers["identitas"]
    col1, col2 = st.columns(2)
    with col1:
        if "nama_notaris" not in st.session_state:
            st.session_state["nama_notaris"] = ident.get("nama_notaris", "")
        st.text_input(
            "Nama Notaris (contoh: Herman Setiawan, S.H., M.Kn)",
            key="nama_notaris"
        )
        ident["nama_notaris"] = st.session_state["nama_notaris"]

        if "username" not in st.session_state:
            st.session_state["username"] = ident.get("username", "")
        st.text_input("Username Akun AHU Online", value=st.session_state["username"], key="username")
        ident["username"] = st.session_state["username"]

        if "alamat" not in st.session_state:
            st.session_state["alamat"] = ident.get("alamat", "")
        st.text_input("Alamat Lengkap Kantor Notaris", value=st.session_state["alamat"], key="alamat")
        ident["alamat"] = st.session_state["alamat"]

        wilayah_idx = daftar_wilayah.index(ident["wilayah_input"]) if ident.get("wilayah_input") in daftar_wilayah else 0
        st.selectbox("Pilih Wilayah Provinsi Kedudukan", daftar_wilayah, index=wilayah_idx, key="wilayah_input")
        ident["wilayah_input"] = st.session_state["wilayah_input"]
    with col2:
        st.text_input("NIK KTP (16 digit angka)", value=ident.get("NIK_KTP", ""), key="NIK_KTP")
        ident["NIK_KTP"] = st.session_state["NIK_KTP"]

        st.text_input("Nomor HP", value=ident.get("nomor_HP", ""), key="nomor_HP")
        ident["nomor_HP"] = st.session_state["nomor_HP"]

        kota_idx = daftar_kota.index(ident["kota"]) if ident.get("kota") in daftar_kota else 0
        st.selectbox("Pilih Kedudukan Kota/Kabupaten", daftar_kota, index=kota_idx, key="kota")
        ident["kota"] = st.session_state["kota"]

    st.markdown('</div>', unsafe_allow_html=True)

    c1, c2 = st.columns([1, 1])
    with c2:
        if st.button("Lanjut", type="primary", use_container_width=True, key="next1"):
            required = [
                ident.get("nama_notaris", ""), ident.get("NIK_KTP", ""), ident.get("username", ""),
                ident.get("nomor_HP", ""), ident.get("alamat", ""), ident.get("kota", ""),
            ]
            nik = ident.get("NIK_KTP", "")
            hp = ident.get("nomor_HP", "")
            if any(f is None or f == "" for f in required):
                st.error("⚠️ Semua data wajib diisi.")
            elif not nik.isdigit() or not hp.isdigit():
                st.error("⚠️ NIK KTP dan Nomor HP harus berupa angka.")
            elif len(nik) != 16:
                st.error("⚠️ NIK KTP harus 16 digit.")
            else:
                go_next()
                st.rerun()

# =========================================================
# STEP 2 — KLIEN SESUAI PROFESI
# =========================================================
elif st.session_state.step == 2:
    st.markdown('<div class="step-card">', unsafe_allow_html=True)
    st.markdown("<h2>2. Jumlah Klien sesuai Profesi</h2>", unsafe_allow_html=True)
    st.markdown('<p class="step-subtext">Isi jumlah klien Tahun 2024–2025 berdasarkan profesi. Boleh dikosongkan (0) bila tidak ada.</p>', unsafe_allow_html=True)

    cols = st.columns(2)
    for i, k in enumerate(profil.keys()):
        with cols[i % 2]:
            st.number_input(k, min_value=0, value=answers["profil"].get(k, 0), key=f"profil_{k}")
            answers["profil"][k] = st.session_state[f"profil_{k}"]

    total_klien = sum(answers["profil"].values())
    st.markdown(f"**Total Jumlah Klien: {total_klien}**")
    st.markdown('</div>', unsafe_allow_html=True)

    c1, c2 = st.columns(2)
    with c1:
        if st.button("Sebelumnya", use_container_width=True, key="prev2"):
            go_prev(); st.rerun()
    with c2:
        if st.button("Lanjut", type="primary", use_container_width=True, key="next2"):
            go_next(); st.rerun()

# =========================================================
# STEP 3 — KLIEN SESUAI BISNIS
# =========================================================
elif st.session_state.step == 3:
    st.markdown('<div class="step-card">', unsafe_allow_html=True)
    st.markdown("<h2>3. Jumlah Klien sesuai Bisnis</h2>", unsafe_allow_html=True)
    st.markdown('<p class="step-subtext">Isi jumlah klien berdasarkan bidang bisnis yang dijalankan.</p>', unsafe_allow_html=True)

    cols = st.columns(2)
    for i, k in enumerate(bisnis_pengguna.keys()):
        with cols[i % 2]:
            st.number_input(k, min_value=0, value=answers["bisnis"].get(k, 0), key=f"bisnis_{k}")
            answers["bisnis"][k] = st.session_state[f"bisnis_{k}"]
    st.markdown('</div>', unsafe_allow_html=True)

    c1, c2 = st.columns(2)
    with c1:
        if st.button("Sebelumnya", use_container_width=True, key="prev3"):
            go_prev(); st.rerun()
    with c2:
        if st.button("Lanjut", type="primary", use_container_width=True, key="next3"):
            go_next(); st.rerun()

# =========================================================
# STEP 4 — JASA & PRODUK
# =========================================================
elif st.session_state.step == 4:
    st.markdown('<div class="step-card">', unsafe_allow_html=True)
    st.markdown("<h2>4. Jasa & Produk yang Diurus Klien</h2>", unsafe_allow_html=True)
    st.markdown('<p class="step-subtext">Bagian A: jenis jasa yang digunakan. Bagian B: dokumen/produk akta yang diurus.</p>', unsafe_allow_html=True)

    st.markdown("**A. Jumlah Klien sesuai Jasa yang Digunakan**")
    cols = st.columns(2)
    for i, k in enumerate(jasa.keys()):
        with cols[i % 2]:
            st.number_input(k, min_value=0, value=answers["jasa"].get(k, 0), key=f"jasa_{k}")
            answers["jasa"][k] = st.session_state[f"jasa_{k}"]

    st.markdown("**B. Jumlah Dokumen/Produk Jasa yang Diurus Klien**")
    cols2 = st.columns(2)
    for i, k in enumerate(produk.keys()):
        with cols2[i % 2]:
            st.number_input(k, min_value=0, value=answers["produk"].get(k, 0), key=f"produk_{k}")
            answers["produk"][k] = st.session_state[f"produk_{k}"]
    st.markdown('</div>', unsafe_allow_html=True)

    c1, c2 = st.columns(2)
    with c1:
        if st.button("Sebelumnya", use_container_width=True, key="prev4"):
            go_prev(); st.rerun()
    with c2:
        if st.button("Lanjut", type="primary", use_container_width=True, key="next4"):
            go_next(); st.rerun()

# =========================================================
# STEP 5 — NEGARA & APGAKKUM
# =========================================================
elif st.session_state.step == 5:
    st.markdown('<div class="step-card">', unsafe_allow_html=True)
    st.markdown("<h2>5. Klien sesuai Negara & Riwayat APGAKKUM</h2>", unsafe_allow_html=True)
    st.markdown('<p class="step-subtext">Isi jumlah klien asal negara tertentu, lalu jawab pertanyaan terkait Aparat Penegak Hukum.</p>', unsafe_allow_html=True)

    cols = st.columns(2)
    for i, k in enumerate(negara.keys()):
        with cols[i % 2]:
            st.number_input(k, min_value=0, value=answers["negara"].get(k, 0), key=f"negara_{k}")
            answers["negara"][k] = st.session_state[f"negara_{k}"]

    st.markdown("**Terkait Aparat Penegak Hukum**")
    apgakkum_opts = ["YA", "TIDAK"]
    apgakkum_idx = apgakkum_opts.index(answers["apgakkum"]) if answers.get("apgakkum") in apgakkum_opts else 1
    st.radio(
        "Apakah Notaris pernah dipanggil atau diminta informasi oleh Aparat Penegak Hukum?",
        apgakkum_opts, index=apgakkum_idx, key="inputs_apgakkum"
    )
    answers["apgakkum"] = st.session_state["inputs_apgakkum"]
    st.markdown('</div>', unsafe_allow_html=True)

    c1, c2 = st.columns(2)
    with c1:
        if st.button("Sebelumnya", use_container_width=True, key="prev5"):
            go_prev(); st.rerun()
    with c2:
        if st.button("Lanjut", type="primary", use_container_width=True, key="next5"):
            go_next(); st.rerun()

# =========================================================
# STEP 6 — KEPATUHAN 1-10
# =========================================================
elif st.session_state.step == 6:
    st.markdown('<div class="step-card">', unsafe_allow_html=True)
    st.markdown("<h2>6. Pertanyaan Kepatuhan Notaris (No. 1–10)</h2>", unsafe_allow_html=True)
    st.markdown('<p class="step-subtext">Jawab sesuai kondisi kepatuhan Kantor Notaris Anda saat ini.</p>', unsafe_allow_html=True)

    qa_opts = ["YA", "TIDAK"]

    def render_q(i):
        idx = qa_opts.index(answers["q"].get(i)) if answers["q"].get(i) in qa_opts else 0
        st.radio(q_full_text[i], qa_opts, index=idx, key=f"q{i}")
        answers["q"][i] = st.session_state[f"q{i}"]

    render_q(1)
    st.file_uploader("Upload Dokumen Pendukung (Form CDD, EDD dan Analisa Resiko) dengan format PDF", type=["pdf"], key="uploaded_file1")
    if st.session_state.get("uploaded_file1") is not None:
        answers["file1"] = st.session_state["uploaded_file1"]
    if answers.get("file1") is not None:
        st.success(f"File berhasil diupload: {answers['file1'].name}")

    render_q(2)
    st.file_uploader("Upload Dokumen Pendukung (SOP PMPJ) dengan format PDF", type=["pdf"], key="uploaded_file2")
    if st.session_state.get("uploaded_file2") is not None:
        answers["file2"] = st.session_state["uploaded_file2"]
    if answers.get("file2") is not None:
        st.success(f"File berhasil diupload: {answers['file2'].name}")

    for i in range(3, 11):
        render_q(i)

    st.markdown('</div>', unsafe_allow_html=True)

    c1, c2 = st.columns(2)
    with c1:
        if st.button("Sebelumnya", use_container_width=True, key="prev6"):
            go_prev(); st.rerun()
    with c2:
        if st.button("Lanjut", type="primary", use_container_width=True, key="next6"):
            if not answers["q"].get(1) or not answers["q"].get(2):
                st.error("⚠️ Pertanyaan No. 1 dan No. 2 wajib dijawab.")
            else:
                go_next(); st.rerun()

# =========================================================
# STEP 7 — KEPATUHAN 11-22
# =========================================================
elif st.session_state.step == 7:
    st.markdown('<div class="step-card">', unsafe_allow_html=True)
    st.markdown("<h2>7. Pertanyaan Kepatuhan Notaris (No. 11–22)</h2>", unsafe_allow_html=True)
    st.markdown('<p class="step-subtext">Lanjutkan menjawab pertanyaan kepatuhan berikut.</p>', unsafe_allow_html=True)

    qa_opts = ["YA", "TIDAK"]
    for i in range(11, 23):
        idx = qa_opts.index(answers["q"].get(i)) if answers["q"].get(i) in qa_opts else 0
        st.radio(q_full_text[i], qa_opts, index=idx, key=f"q{i}")
        answers["q"][i] = st.session_state[f"q{i}"]

    st.markdown('</div>', unsafe_allow_html=True)

    c1, c2 = st.columns(2)
    with c1:
        if st.button("Sebelumnya", use_container_width=True, key="prev7"):
            go_prev(); st.rerun()
    with c2:
        if st.button("Lanjut", type="primary", use_container_width=True, key="next7"):
            go_next(); st.rerun()

# =========================================================
# STEP 8 — KEPATUHAN 23-34
# =========================================================
elif st.session_state.step == 8:
    st.markdown('<div class="step-card">', unsafe_allow_html=True)
    st.markdown("<h2>8. Pertanyaan Kepatuhan Notaris (No. 23–34)</h2>", unsafe_allow_html=True)
    st.markdown('<p class="step-subtext">Bagian terakhir dari pertanyaan kepatuhan.</p>', unsafe_allow_html=True)

    qa_opts = ["YA", "TIDAK"]
    for i in range(23, 35):
        idx = qa_opts.index(answers["q"].get(i)) if answers["q"].get(i) in qa_opts else 0
        st.radio(q_full_text[i], qa_opts, index=idx, key=f"q{i}")
        answers["q"][i] = st.session_state[f"q{i}"]

    st.markdown('</div>', unsafe_allow_html=True)

    c1, c2 = st.columns(2)
    with c1:
        if st.button("Sebelumnya", use_container_width=True, key="prev8"):
            go_prev(); st.rerun()
    with c2:
        if st.button("Lanjut ke Ringkasan", type="primary", use_container_width=True, key="next8"):
            if not answers["q"].get(34):
                st.error("⚠️ Pertanyaan No. 34 wajib dijawab.")
            else:
                go_next(); st.rerun()

# =========================================================
# STEP 9 — RINGKASAN & KIRIM
# =========================================================
    st.write("DEBUG IDENT")
    st.write(answers["identitas"])
elif st.session_state.step == 9:
    st.markdown('<div class="step-card">', unsafe_allow_html=True)
    st.markdown("<h2>9. Ringkasan & Kirim</h2>", unsafe_allow_html=True)
    st.markdown('<p class="step-subtext">Periksa kembali ringkasan berikut sebelum mengirim kuisioner.</p>', unsafe_allow_html=True)

    # DIPERBAIKI: semua nilai diambil dari dictionary permanen `answers`,
    # BUKAN dari st.session_state["key_widget"] langsung, karena key widget
    # dari step-step sebelumnya sudah tidak ada lagi begitu step berpindah.
    ident = answers["identitas"]
    inputs_profil = answers["profil"]
    inputs_bisnis = answers["bisnis"]
    inputs_jasa = answers["jasa"]
    inputs_produk = answers["produk"]
    inputs_negara = answers["negara"]
    inputs_apgakkum = answers["apgakkum"]
    jumlah_klien = sum(inputs_profil.values())
    wilayah_input = ident.get("wilayah_input")

    colA, colB, colC = st.columns(3)
    with colA:
        st.metric("Nama Notaris", ident.get("nama_notaris", "-") or "-")
    with colB:
        st.metric("Total Jumlah Klien", jumlah_klien)
    with colC:
        st.metric("Wilayah", wilayah_input or "-")

    with st.expander("Lihat rincian identitas & data yang sudah diisi"):
            st.markdown('<div class="report-block"><div class="report-title">Identitas</div>', unsafe_allow_html=True)
            st.markdown(render_report_table([
                ("Nama Notaris", ident.get("nama_notaris")),
                ("NIK KTP", ident.get("NIK_KTP")),
                ("Username AHU Online", ident.get("username")),
                ("Nomor HP", ident.get("nomor_HP")),
                ("Alamat", ident.get("alamat")),
                ("Kota/Kabupaten", ident.get("kota")),
                ("Wilayah", wilayah_input),
            ]), unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)

            # Hanya tampilkan item dengan jumlah > 0 agar laporan ringkas & rapi
            def _nonzero_rows(d):
                return [(k, v) for k, v in d.items() if v] or [("(Tidak ada data)", "")]

            st.markdown('<div class="report-block"><div class="report-title">Jumlah Klien sesuai Profesi</div>', unsafe_allow_html=True)
            st.markdown(render_report_table(_nonzero_rows(inputs_profil)), unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)

            st.markdown('<div class="report-block"><div class="report-title">Jumlah Klien sesuai Bisnis</div>', unsafe_allow_html=True)
            st.markdown(render_report_table(_nonzero_rows(inputs_bisnis)), unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)

            st.markdown('<div class="report-block"><div class="report-title">Jumlah Klien sesuai Jasa</div>', unsafe_allow_html=True)
            st.markdown(render_report_table(_nonzero_rows(inputs_jasa)), unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)

            st.markdown('<div class="report-block"><div class="report-title">Jumlah Dokumen/Produk</div>', unsafe_allow_html=True)
            st.markdown(render_report_table(_nonzero_rows(inputs_produk)), unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)

            st.markdown('<div class="report-block"><div class="report-title">Jumlah Klien sesuai Negara</div>', unsafe_allow_html=True)
            st.markdown(render_report_table(_nonzero_rows(inputs_negara)), unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)

            st.markdown('<div class="report-block"><div class="report-title">APGAKKUM</div>', unsafe_allow_html=True)
            st.markdown(render_report_table([
                ("Pernah dipanggil/diminta informasi oleh APGAKKUM", inputs_apgakkum),
            ]), unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)

            st.markdown('<div class="report-block"><div class="report-title">Jawaban Kepatuhan (No. 1-34)</div>', unsafe_allow_html=True)
            st.markdown(render_report_table([
                (q_full_text[i][:80] + ("…" if len(q_full_text[i]) > 80 else ""), answers["q"].get(i))
                for i in range(1, 35)
            ]), unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)

            st.markdown('<div class="report-block"><div class="report-title">Dokumen Pendukung</div>', unsafe_allow_html=True)
            st.markdown(render_report_table([
                ("Dokumen Q1 (CDD/EDD/Analisis Risiko)", answers["file1"].name if answers.get("file1") else "Belum ada"),
                ("Dokumen Q2 (SOP PMPJ)", answers["file2"].name if answers.get("file2") else "Belum ada"),
            ]), unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)

    c1, c2 = st.columns(2)
    with c1:
        if st.button("Sebelumnya", use_container_width=True, key="prev9"):
            go_prev(); st.rerun()
    with c2:
        submitted = st.button("Kirim Kuisioner", type="primary", use_container_width=True, key="submit_final")

    if submitted:
        nama_notaris = ident.get("nama_notaris", "")
        NIK_KTP = ident.get("NIK_KTP", "")
        username = ident.get("username", "")
        nomor_HP = ident.get("nomor_HP", "")
        alamat = ident.get("alamat", "")
        kota = ident.get("kota", "")
        q_list = [answers["q"].get(i) for i in range(1, 35)]
        q1 = q_list[0]
        uploaded_file1 = answers.get("file1")
        uploaded_file2 = answers.get("file2")

        required_fields = [nama_notaris, NIK_KTP, username, nomor_HP, alamat, kota, q_list[0], q_list[1], q_list[33]]
        missing = any(f is None or f == "" for f in required_fields)

        if missing:
            st.error("⚠️ Semua data wajib diisi (kecuali dokumen pendukung).")
        elif not NIK_KTP.isdigit() or not nomor_HP.isdigit():
            st.error("⚠️ NIK KTP dan Nomor HP harus berupa angka.")
        elif len(NIK_KTP) != 16:
            st.error("⚠️ NIK KTP harus 16 digit.")
        else:
            kata_kunci_list = [
                "Formulir Customer Due Diligence",
                "formulir customer due diligence perorangan",
                "Analisis Risiko", "Analisis Resiko",
                "Enhanced Due Diligence",
                "CDD",
                "EDD"]
            is_valid_ocr_q1, teks_ocr_q1, jumlah_kata_ditemukan_q1 = validasi_ocr_pdf(
                uploaded_file1, kata_kunci_list, judul="Dokumen Q1 (CDD/EDD/Analisis Risiko)"
            )

            os.makedirs("uploads", exist_ok=True)
            doc1_path, doc2_path = "", ""
            file_link_1, file_link_2 = "", ""

            def upload_to_drive(local_path, original_name):
                if drive_service is None:
                    return local_path
                try:
                    FOLDER_ID = "1v0HSHab3hTRLPBDX4Sk5SzfHay2-rG8N"
                    file_metadata = {"name": original_name, "parents": [FOLDER_ID]}
                    media = MediaFileUpload(local_path, mimetype="application/pdf")
                    uploaded = drive_service.files().create(body=file_metadata, media_body=media, fields="id").execute()
                    file_id = uploaded.get("id")
                    drive_service.permissions().create(fileId=file_id, body={"type": "anyone", "role": "reader"}).execute()
                    return f"https://drive.google.com/file/d/{file_id}/view?usp=sharing"
                except Exception as e:
                    st.error(f"❌ Gagal upload ke Google Drive: {e}")
                    return local_path

            if uploaded_file1 is not None:
                filename_1 = f"{datetime.now().strftime('%Y%m%d_%H%M%S')}_doc1_{uploaded_file1.name}"
                doc1_path = os.path.join("uploads", filename_1)
                with open(doc1_path, "wb") as f:
                    f.write(uploaded_file1.getbuffer())
                file_link_1 = upload_to_drive(doc1_path, uploaded_file1.name)

            if uploaded_file2 is not None:
                filename_2 = f"{datetime.now().strftime('%Y%m%d_%H%M%S')}_doc2_{uploaded_file2.name}"
                doc2_path = os.path.join("uploads", filename_2)
                with open(doc2_path, "wb") as f:
                    f.write(uploaded_file2.getbuffer())
                file_link_2 = upload_to_drive(doc2_path, uploaded_file2.name)

            hasil_inherent = hitung_risiko({
                "profil": inputs_profil, "bisnis": inputs_bisnis, "jasa": inputs_jasa,
                "negara": inputs_negara, "apgakkum": inputs_apgakkum, "wilayah": wilayah_input
            })
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            q_answers = {i: answers["q"].get(i, "TIDAK") for i in range(1, 34)}
            nilai_ic, kategori_ic = hitung_internal_control(q_answers, uploaded_file1, is_valid_ocr_q1)
            kategori_residual, nilai_residual = hitung_residual_risk(hasil_inherent["kategori_risiko"], kategori_ic)
            nilai_pengguna, kategori_pengguna = risiko_pengguna_jasa(jumlah_klien)

            df_temp = pd.DataFrame([{"Nilai Residual Risk": nilai_residual, "Nilai Risiko Pengguna Jasa": nilai_pengguna}])
            kategori_final = final_risk(df_temp).loc[0, "Tingkat Risiko"]

            data = {
                "Timestamp": timestamp,
                "Nama Notaris": nama_notaris.title(),
                "NIK KTP": NIK_KTP,
                "Username Akun AHU Online": username,
                "Nomor HP": nomor_HP,
                "2. Alamat Lengkap Kantor Notaris": alamat,
                "Kedudukan Kota/Kabupaten": kota,
                "3. Jumlah Klien Tahun 2024-2025": jumlah_klien,
                "Wilayah": wilayah_input
            }
            data.update({k: inputs_profil.get(k, 0) for k in profil.keys()})
            data.update({k: inputs_bisnis.get(k, 0) for k in bisnis_pengguna.keys()})
            data.update({k: inputs_jasa.get(k, 0) for k in jasa.keys()})
            data.update({k: inputs_produk.get(k, 0) for k in produk.keys()})
            data.update({k: inputs_negara.get(k, 0) for k in negara.keys()})
            data.update({k: wilayah_skor.get(k, 0) for k in wilayah_skor.keys()})

            for i, col in enumerate(q_cols):
                data[col] = q_list[i]
            data["Dokumen_Pendukung (Q1)"] = file_link_1
            data["Dokumen Pendukung (SOP PMPJ) (Q2)"] = file_link_2

            data["Apakah Notaris pernah dipanggil atau diminta informasi oleh Aparat Penegak Hukum?"] = inputs_apgakkum
            data["jawaban_profil"] = hasil_inherent["jawaban_profil"]
            data["skor_profil"] = hasil_inherent["skor_profil"]
            data["jawaban_bisnis"] = hasil_inherent["jawaban_bisnis"]
            data["skor_bisnis"] = hasil_inherent["skor_bisnis"]
            data["jawaban_jasa"] = hasil_inherent["jawaban_jasa"]
            data["wilayah"] = wilayah_input
            data["skor_jasa"] = hasil_inherent["skor_jasa"]
            data["jawaban_negara"] = hasil_inherent["jawaban_negara"]
            data["skor_negara"] = hasil_inherent["skor_negara"]
            data["jawaban_apgakkum"] = hasil_inherent["jawaban_apgakkum"]
            data["skor_apgakkum"] = hasil_inherent["skor_apgakkum"]
            data["jawaban_wilayah"] = hasil_inherent["jawaban_wilayah"]
            data["skor_wilayah"] = hasil_inherent["skor_wilayah"]

            data["Nilai Inherent Risk"] = hasil_inherent["total_skor"]
            data["Tingkat Inherent Risk"] = hasil_inherent["kategori_risiko"]
            data["Nilai Internal Control"] = nilai_ic
            data["Tingkat Internal Control"] = kategori_ic
            data["Tingkat Residual Risk"] = kategori_residual
            data["Nilai Residual Risk"] = nilai_residual
            data["Nilai Risiko Pengguna Jasa"] = nilai_pengguna
            data["Tingkat Risiko Pengguna Jasa"] = kategori_pengguna
            data["Tingkat Risiko"] = kategori_final

            ident_cols = [
                "Timestamp", "Nama Notaris", "NIK KTP", "Username Akun AHU Online", "Nomor HP", "Wilayah",
                "2. Alamat Lengkap Kantor Notaris", "Kedudukan Kota/Kabupaten", "3. Jumlah Klien Tahun 2024-2025"
            ]
            ringkasan_cols = [
                "Nilai Inherent Risk", "Tingkat Inherent Risk",
                "Nilai Internal Control", "Tingkat Internal Control",
                "Tingkat Residual Risk", "Nilai Residual Risk",
                "Nilai Risiko Pengguna Jasa", "Tingkat Risiko Pengguna Jasa", "Tingkat Risiko"
            ]
            q_cols_with_docs = q_cols + ["Dokumen_Pendukung (Q1)", "Dokumen Pendukung (SOP PMPJ) (Q2)"]
            detail_cols = list(profil.keys()) + list(bisnis_pengguna.keys()) + list(jasa.keys()) + list(produk.keys()) + list(negara.keys())
            pilihan_cols = [
                "Apakah Notaris pernah dipanggil atau diminta informasi oleh Aparat Penegak Hukum?",
                "jawaban_profil", "skor_profil", "jawaban_bisnis", "skor_bisnis",
                "jawaban_jasa", "skor_jasa", "jawaban_negara", "skor_negara", "skor_wilayah",
                "jawaban_apgakkum", "skor_apgakkum"
            ]
            column_order = ident_cols + detail_cols + q_cols_with_docs + pilihan_cols + ringkasan_cols

            # =========================================================
            # Helper: Convert nomor kolom ke huruf Excel
            # =========================================================
            def colnum_to_excel(n: int) -> str:
                result = ""
                while n > 0:
                    n, remainder = divmod(n - 1, 26)
                    result = chr(65 + remainder) + result
                return result


            if client is None:
                st.error("❌ Silahkan refresh laman ini dan lakukan pengisian ulang.")

            else:
                SPREADSHEET_NAME = "Kuisioner PMPJ Notaris FINAL 2026"
                SPREADSHEET_ID = "110TjnLF8T_rtK3Y_h_Lc3fRiAT5_Zhrz9fT1ycBnpD0"

                try:

                    # ==========================
                    # Buka Spreadsheet
                    # ==========================
                    try:
                        sh2 = client.open_by_key(SPREADSHEET_ID)
                    except gspread.SpreadsheetNotFound:
                        sh2 = client.open(SPREADSHEET_NAME)

                    worksheet2 = sh2.sheet1

                    # ==========================
                    # Pastikan header ada
                    # ==========================
                    existing_header = worksheet2.row_values(1)

                    if not existing_header:
                        worksheet2.append_row(
                            column_order,
                            value_input_option="RAW"
                        )

                    nama_baru = data.get("Nama Notaris", "")
                    nik_baru = str(data.get("NIK KTP", "")).strip()

                    row_values = [
                        str(data.get(col, ""))
                        for col in column_order
                    ]

                    # ==========================
                    # Cari apakah NIK sudah ada
                    # ==========================
                    nik_col_idx = column_order.index("NIK KTP") + 1

                    target_row = None

                    if nik_baru:

                        nik_column_values = worksheet2.col_values(nik_col_idx)

                        for idx, val in enumerate(nik_column_values):

                            # Skip header
                            if idx == 0:
                                continue

                            if str(val).strip() == nik_baru:
                                target_row = idx + 1
                                break

                    # ==========================
                    # Update atau Append
                    # ==========================
                    if target_row:

                        last_col = colnum_to_excel(len(column_order))

                        worksheet2.update(
                            f"A{target_row}:{last_col}{target_row}",
                            [row_values]
                        )

                        st.warning(
                            f"⚠️ Data lama untuk '{nama_baru}' "
                            f"(NIK: {nik_baru}) ditemukan dan telah diperbarui."
                        )

                    else:

                        worksheet2.append_row(
                            row_values,
                            value_input_option="RAW"
                        )

                        st.info(
                            f"✅ Data baru untuk '{nama_baru}' ditambahkan."
                        )

                    st.success(
                        "✅ Data berhasil disimpan. "
                        "Silakan screenshot laman ini sebagai bukti telah melakukan pengisian kuisioner PMPJ. TIDAK DISARANKAN untuk klik 'kirim kuisioner' lagi. Jika ingin memperbarui isi kuisiner, silahkan refresh laman ini!"
                    )

                    # Reset wizard
                    st.session_state.pop("answers", None)

                except Exception:
                    import traceback

                    st.error(
                        f"❌ Klik sekali lagi untuk menyimpan kuisioner. Error saat menyimpan:\n{traceback.format_exc()}"
                    )

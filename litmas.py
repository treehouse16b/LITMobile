from re import S, sub
import streamlit as st
from docxtpl import DocxTemplate
import io
import os
import pandas as pd
from datetime import datetime
import json
import openpyxl
import matplotlib.pyplot as plt
import google.generativeai as genai

# ==========================================
# KONFIGURASI GEMINI (SDK terbaru)
# ==========================================
GEMINI_API_KEY = st.secrets.get("GEMINI_API_KEY")

if not GEMINI_API_KEY:
    st.error("API Key Gemini belum diset di .streamlit/secrets.toml")
    st.stop()

# Inisialisasi client & model global
genai.configure(api_key=GEMINI_API_KEY)
model = genai.GenerativeModel('gemini-1.5-flash')

st.set_page_config(page_title="LITMAS Bapas Balikpapan + AI Gemini", layout="wide")


# ==========================================
# FUNGSI BANTUAN
# ==========================================
def format_tgl_indo(tgl_obj):
    if tgl_obj is None or pd.isna(tgl_obj):
        return ""
    if isinstance(tgl_obj, str):
        try:
            tgl_obj = datetime.strptime(tgl_obj, "%d/%m/%Y")
        except:
            return tgl_obj
    return tgl_obj.strftime("%d/%m/%Y")


def format_tgl_tulis_indo(tgl_obj):
    if tgl_obj is None or pd.isna(tgl_obj):
        return ""
    if isinstance(tgl_obj, str):
        try:
            tgl_obj = datetime.strptime(tgl_obj, "%d/%m/%Y")
        except:
            return tgl_obj
    bulan = [
        "Januari", "Februari", "Maret", "April", "Mei", "Juni",
        "Juli", "Agustus", "September", "Oktober", "November", "Desember"
    ]
    return f"{tgl_obj.day} {bulan[tgl_obj.month - 1]} {tgl_obj.year}"


@st.cache_data(ttl=300)
def generate_narasi_ai(kategori, d):
    if kategori == "hukum":
        return (
            f"Berdasarkan data yuridis, Klien a.n. {d.get('nama_klien', '...')} terjerat perkara {d.get('perkara', '...')} "
            f"sebagaimana diatur dalam Pasal {d.get('pasal', '...')} dengan vonis {d.get('vonis', '...')}. "
            f"Tindak pidana tersebut dilatarbelakangi oleh faktor {d.get('latarbelakang_tp', '...')}. "
            f"Adapun akibat yang ditimbulkan bagi korban/masyarakat adalah {d.get('akibat_masyarakat', '...')}. "
            f"Saat ini, klien menunjukkan sikap {d.get('sikap_klien_bab7', '...')}, yang mengindikasikan proses perubahan perilaku."
        )

    elif kategori == "klien":
        return (
            f"Klien merupakan seorang {d.get('jk_klien')} bersuku {d.get('suku')} dengan pendidikan terakhir {d.get('pend_terakhir')}. "
            f"Sebelum menjalani pidana, klien bekerja sebagai {d.get('kerja_terakhir')}. "
            f"Terkait pola hidup, klien memiliki riwayat konsumsi sebagai berikut: Rokok ({d.get('riwayat_rokok')}), "
            f"Alkohol ({d.get('riwayat_alkohol')}), dan NAPZA ({d.get('riwayat_napza')}). "
            f"Data ini menjadi indikator penting dalam penyusunan program pembimbingan kepribadian."
        )

    elif kategori == "penjamin":
        return (
            f"Penjamin atas nama {d.get('nama_pjm')} memiliki hubungan sebagai {d.get('hub_pjm')} dengan klien. "
            f"Penjamin menyatakan kesanggupan dan kelayakan yang dibuktikan dengan: {d.get('layak_pjm')}. "
            f"Kondisi sosial dan kesiapan keluarga ini diproyeksikan dapat menjadi faktor pendukung yang efektif selama masa reintegrasi sosial."
        )

    elif kategori == "lingkungan":
        return (
            f"Klien telah mengikuti pembinaan kepribadian berupa {d.get('bina_pribadi')} dan kemandirian berupa {d.get('bina_mandiri')}. "
            f"Potensi minat dan bakat klien di bidang {d.get('minat_bakat')} dapat dikembangkan lebih lanjut. "
            f"Sementara itu, kondisi lingkungan tempat tinggal penjamin digambarkan sebagai berikut: {d.get('kondisi_alam')}. "
            f"Penerimaan masyarakat setempat ({d.get('relasi_masyarakat')}) menjadi modal sosial positif."
        )

    elif kategori == "rel_kel":
        return "Relasi sosial keluarga klien tergolong baik dengan tingkat dukungan yang memadai untuk proses reintegrasi sosial."

    elif kategori == "kronologi":
        return f"Kronologi tindak pidana dimulai dari {d.get('kronologi_detail', '...')}. Keadaan korban pada saat kejadian adalah {d.get('keadaan_korban', '...')}."

    elif kategori == "akibat":
        return f"Akibat yang ditimbulkan terhadap klien adalah {d.get('akibat_klien', '...')}, sementara bagi keluarga dan masyarakat adalah {d.get('akibat_masyarakat', '...')}."

    elif kategori == "sikap_bab7":
        return f"Sikap dan tanggapan klien terhadap proses pemidanaan menunjukkan {d.get('sikap_klien_bab7', '...')}, yang menjadi indikator penting dalam pembimbingan."

    elif kategori == "hasil_rekomendasi":
        return f"Hasil asesmen menunjukkan risiko {d.get('kat_rri', 'RENDAH')} dan kebutuhan {d.get('kat_krim', 'RENDAH')}. Direkomendasikan {d.get('program', 'Pembebasan Bersyarat')} dengan pembinaan intensif pada faktor dominan."

    elif kategori == "kesimpulan":
        return "Kesimpulan: Klien memiliki potensi reintegrasi sosial yang baik dengan dukungan keluarga dan lingkungan. Direkomendasikan pemberian hak bersyarat dengan pengawasan ketat."

    return ""


def sync_penjamin_data(ctx, pilihan_pjm):
    if pilihan_pjm == "Ayah":
        for field in ["nama", "ttl", "agama", "suku", "pend", "kerja", "alamat", "no_hp"]:
            ctx[f"{field}_pjm"] = ctx.get(f"{field}_ayah", "")
        ctx["hub_pjm"] = "Ayah Kandung"
    elif pilihan_pjm == "Ibu":
        for field in ["nama", "ttl", "agama", "suku", "pend", "kerja", "alamat", "no_hp"]:
            ctx[f"{field}_pjm"] = ctx.get(f"{field}_ibu", "")
        ctx["hub_pjm"] = "Ibu Kandung"
    return ctx


def validate_data(d):
    required = ["nama_klien", "no_reg_litmas", "perkara", "pasal", "vonis"]
    missing = [k for k in required if not d.get(k)]
    if missing:
        st.warning(f"Field wajib diisi: {', '.join(missing)}")
        return False
    return True

# ==========================================
# FUNGSI ASSESMENT KRIMINOGENIK (dari model_kriminogenik.py)
# ==========================================
def hitung_skor_domain(responses):
    """
    responses: dict dengan key 1-29, value skor numerik
    Return: dict skor per domain + total
    """
    skor = {
        "keluarga_pernikahan": responses.get(1, 0) + responses.get(2, 0) + responses.get(3, 0),
        "pendidikan_pekerjaan": (
            responses.get(4, 0) + responses.get(5, 0) + responses.get(6, 0) + responses.get(7, 0) +
            responses.get(8, 0) + responses.get(9, 0) + responses.get(10, 0) + responses.get(11, 0)
        ),
        "narkotika_alkohol": (
            responses.get(12, 0) + responses.get(13, 0) + responses.get(14, 0) +
            responses.get(15, 0) + responses.get(16, 0) + responses.get(17, 0)
        ),
        "hubungan_sosial": responses.get(18, 0) + responses.get(19, 0) + responses.get(20, 0),
        "waktu_luang": responses.get(21, 0) + responses.get(22, 0),
        "manajemen_keuangan": responses.get(23, 0) + responses.get(24, 0),
        "sikap_antisosial": (
            responses.get(25, 0) + responses.get(26, 0) + responses.get(27, 0) +
            responses.get(28, 0) + responses.get(29, 0)
        ),
    }
    total = sum(skor.values())
    return skor, total


def kategori_domain(skor_domain):
    kat = {}
    ranges = {
        "keluarga_pernikahan": [(0,1,"Rendah","green"), (2,3,"Sedang","orange"), (4,5,"Tinggi","red"), (6,6,"Sangat Tinggi","darkred")],
        "pendidikan_pekerjaan": [(0,2,"Rendah","green"), (3,6,"Sedang","orange"), (7,8,"Tinggi","red"), (9,10,"Sangat Tinggi","darkred")],
        "narkotika_alkohol": [(0,1,"Rendah","green"), (2,4,"Sedang","orange"), (5,5,"Tinggi","red"), (6,6,"Sangat Tinggi","darkred")],
        "hubungan_sosial": [(0,1,"Rendah","green"), (2,3,"Sedang","orange"), (4,4,"Tinggi","red"), (5,5,"Sangat Tinggi","darkred")],
        "waktu_luang": [(0,0,"Rendah","green"), (1,1,"Sedang","orange"), (2,2,"Tinggi","red")],
        "manajemen_keuangan": [(0,0,"Rendah","green"), (1,1,"Sedang","orange"), (2,2,"Tinggi","red")],
        "sikap_antisosial": [(0,1,"Rendah","green"), (2,4,"Sedang","orange"), (5,6,"Tinggi","red"), (7,7,"Sangat Tinggi","darkred")]
    }
    for domain, rg in ranges.items():
        skor = skor_domain.get(domain, 0)
        for low, high, kat_name, warna in rg:
            if low <= skor <= high:
                kat[domain] = (kat_name, warna)
                break
        else:
            kat[domain] = ("Tidak Terdefinisi", "gray")
    return kat


def kategori_total(total):
    if total <= 10:
        return "Rendah", "green"
    elif total <= 22:
        return "Sedang", "orange"
    elif total <= 29:
        return "Tinggi", "red"
    else:
        return "Sangat Tinggi", "darkred"


def get_faktor_dominan(skor_domain):
    max_skor = max(skor_domain.values())
    dominan = [d for d, s in skor_domain.items() if s == max_skor]
    return ", ".join(d.capitalize().replace("_", " ") for d in dominan), max_skor


def get_kesimpulan_dan_rekomendasi(total, kategori_total, skor_domain, faktor_dominan, centang_b1, centang_b2):
    kesimpulan = f"Total skor kebutuhan kriminogenik: **{total}** → Kategori **{kategori_total}**.\n"
    kesimpulan += f"Faktor dominan: **{faktor_dominan}** (skor tertinggi).\n"
    
    if centang_b1:
        kesimpulan += f"Ada {len(centang_b1)} pertimbangan tindak pidana tertentu.\n"
    if centang_b2:
        kesimpulan += f"Ada {len(centang_b2)} faktor kebutuhan tambahan.\n"

    rekomendasi = "**Rekomendasi Pembinaan (berdasarkan prinsip RNR):**\n"
    
    if "Rendah" in kategori_total:
        rekomendasi += "- Pembinaan standar: fokus pada penguatan dukungan sosial dan kemandirian.\n"
        rekomendasi += "- Pantau faktor dominan agar tidak meningkat.\n"
    elif "Sedang" in kategori_total:
        rekomendasi += "- Susun case plan individual berdasarkan faktor dominan.\n"
        rekomendasi += "- Program sedang: konseling, pelatihan keterampilan, dukungan keluarga.\n"
    elif "Tinggi" in kategori_total:
        rekomendasi += "- Intervensi intensif: prioritas pada domain dominan.\n"
        rekomendasi += "- Konseling psikologis, rehabilitasi (jika narkoba/alkohol tinggi), manajemen emosi.\n"
    else:
        rekomendasi += "- Intervensi maksimal: asesmen mendalam + spesialis.\n"
        rekomendasi += "- Program khusus: terapi perilaku, rehabilitasi NAPZA, pengawasan ketat.\n"

    if skor_domain.get("narkotika_alkohol", 0) >= 5:
        rekomendasi += "- Prioritaskan rehabilitasi NAPZA dan pencegahan relapse.\n"
    if skor_domain.get("keluarga_pernikahan", 0) >= 4:
        rekomendasi += "- Konseling keluarga dan perbaikan hubungan interpersonal.\n"
    if skor_domain.get("pendidikan_pekerjaan", 0) >= 7:
        rekomendasi += "- Pelatihan kerja, pendidikan lanjutan, dan kemandirian ekonomi.\n"
    if skor_domain.get("sikap_antisosial", 0) >= 5:
        rekomendasi += "- Program pengembangan empati, tanggung jawab, dan sikap pro-sosial.\n"

    rekomendasi += "\n**Catatan:** Integrasikan dengan hasil RRI. Validasi oleh petugas profesional."

    return kesimpulan, rekomendasi

# ==========================================
# INISIALISASI DATA
# ==========================================
if "data" not in st.session_state:
    st.session_state.data = {
        "upt": "",
        "no_reg_litmas": "",
        "no_surat": "",
        "tgl_surat": datetime.now(),
        "perihal": "",
        "program": "Pembebasan Bersyarat",
        "tgl_tpp": datetime.now(),
        "tgl_1per3": datetime.now(),
        "tgl_1per2": datetime.now(),
        "tgl_2per3": datetime.now(),
        "ekspirasi": "",
        "perkara": "",
        "pasal": "",
        "vonis": "",
        "latarbelakang_tp": "",
        "kronologi_detail": "",
        "keadaan_korban": "",
        "akibat_klien": "",
        "akibat_masyarakat": "",
        "sikap_klien_bab7": "",
        "nama_klien": "",
        "no_reg_bin": "",
        "ttl_klien": "",
        "jk_klien": "Laki-laki",
        "suku": "",
        "warga_negara": "Indonesia",
        "agama": "",
        "pend_terakhir": "",
        "kerja_terakhir": "",
        "alamat_klien": "",
        "ciri_khusus": "",
        "status_nikah": "",
        "riwayat_nikah_klien": "",
        "sd": "",
        "smp": "",
        "sma": "",
        "pt": "",
        "riwayat_rokok": "",
        "riwayat_alkohol": "",
        "riwayat_napza": "",
        "nama_ayah": "",
        "ttl_ayah": "",
        "agama_ayah": "",
        "suku_ayah": "",
        "pend_ayah": "",
        "kerja_ayah": "",
        "alamat_ayah": "",
        "no_hp_ayah": "",
        "ket_ayah": "",
        "nama_ibu": "",
        "ttl_ibu": "",
        "agama_ibu": "",
        "suku_ibu": "",
        "pend_ibu": "",
        "kerja_ibu": "",
        "alamat_ibu": "",
        "no_hp_ibu": "",
        "ket_ibu": "",
        "nama_pjm": "",
        "ttl_pjm": "",
        "agama_pjm": "",
        "suku_pjm": "",
        "pend_pjm": "",
        "kerja_pjm": "",
        "alamat_pjm": "",
        "hub_pjm": "",
        "no_hp_pjm": "",
        "layak_pjm": "",
        "riwayat_nikah_pjm": "",
        "saudara_data": pd.DataFrame(
            columns=["No", "Nama", "Usia", "JK", "Pendidikan", "Pekerjaan", "Keterangan"],
            data=[
                {"No": 1, "Nama": "", "Usia": "", "JK": "L", "Pendidikan": "", "Pekerjaan": "", "Keterangan": ""},
                {"No": 2, "Nama": "", "Usia": "", "JK": "L", "Pendidikan": "", "Pekerjaan": "", "Keterangan": ""},
                {"No": 3, "Nama": "", "Usia": "", "JK": "L", "Pendidikan": "", "Pekerjaan": "", "Keterangan": ""},
            ]
        ),
        "bina_pribadi": "",
        "bina_mandiri": "",
        "minat_bakat": "",
        "relasi_masyarakat": "",
        "kondisi_alam": "",
        "relasi_kel": "",
        "skor_rri": 0,
        "kat_rri": "RENDAH",
        "skor_krim": 0,
        "kat_krim": "RENDAH",
        "ai_analisis_hukum": "",
        "ai_analisis_klien": "",
        "ai_analisis_penjamin": "",
        "ai_analisis_lingkungan": "",
        "gender": "Laki-laki",
        "kasus_narkotika": False,
    }


# Inisialisasi halaman default
if "halaman" not in st.session_state:
    st.session_state.halaman = "1. Administrasi & Hukum"


# ==========================================
# SIDEBAR NAVIGASI (selalu muncul)
# ==========================================
with st.sidebar:
    st.title("📂 Menu Litmas")
    pilihan = [
        "1. Administrasi & Hukum",
        "2. Identitas Klien",
        "3. Keluarga (Ayah & Ibu)",
        "4. Penjamin & Saudara",
        "5. Riwayat & Lingkungan",
        "6. Asesmen (RRI)",
        "6. Asesmen Faktor Kriminogenik",
        "7. Download",
    ]
    index_default = pilihan.index(st.session_state.halaman) if st.session_state.halaman in pilihan else 0
    halaman_pilihan = st.radio("Pilih Bagian", pilihan, index=index_default, key="nav_radio")
    st.session_state.halaman = halaman_pilihan


# ==========================================
# KONTEN UTAMA
# ==========================================
halaman = st.session_state.halaman

if halaman == "1. Administrasi & Hukum":
    st.header("🏢 Administrasi & Hukum")

    c1, c2, c3 = st.columns(3)
    st.session_state.data["upt"] = c1.text_input("Asal UPT", st.session_state.data["upt"])
    st.session_state.data["no_reg_litmas"] = c2.text_input("No. Reg Litmas", st.session_state.data["no_reg_litmas"])
    st.session_state.data["no_surat"] = c3.text_input("No. Surat Permintaan", st.session_state.data["no_surat"])

    c1, c2, c3 = st.columns(3)
    st.session_state.data["tgl_surat"] = c1.date_input("Tanggal Surat", st.session_state.data["tgl_surat"], format="DD/MM/YYYY")
    st.session_state.data["perihal"] = c2.text_input("Perihal", st.session_state.data["perihal"])
    st.session_state.data["program"] = c3.selectbox(
        "Program Reintegrasi",
        ["Pembebasan Bersyarat", "Cuti Bersyarat", "Asimilasi"],
        index=["Pembebasan Bersyarat", "Cuti Bersyarat", "Asimilasi"].index(st.session_state.data["program"])
    )

    st.subheader("Data Tahapan & Putusan")
    c1, c2, c3, c4 = st.columns(4)
    st.session_state.data["perkara"] = c1.text_input("Perkara", st.session_state.data["perkara"])
    st.session_state.data["pasal"] = c2.text_input("Pasal", st.session_state.data["pasal"])
    st.session_state.data["vonis"] = c3.text_input("Vonis", st.session_state.data["vonis"])
    st.session_state.data["ekspirasi"] = c4.text_input("Tgl Ekspirasi", st.session_state.data["ekspirasi"])

    c1, c2, c3, c4 = st.columns(4)
    st.session_state.data["tgl_tpp"] = c1.date_input("Tgl Sidang TPP", st.session_state.data["tgl_tpp"], format="DD/MM/YYYY")
    st.session_state.data["tgl_1per3"] = c2.date_input("Tgl 1/3", st.session_state.data["tgl_1per3"], format="DD/MM/YYYY")
    st.session_state.data["tgl_1per2"] = c3.date_input("Tgl 1/2", st.session_state.data["tgl_1per2"], format="DD/MM/YYYY")
    st.session_state.data["tgl_2per3"] = c4.date_input("Tgl 2/3", st.session_state.data["tgl_2per3"], format="DD/MM/YYYY")

    st.subheader("⚖️ Data Hukum (Bab VI & VII)")
    st.session_state.data["latarbelakang_tp"] = st.text_area("Latar Belakang Tindak Pidana", st.session_state.data["latarbelakang_tp"])
    st.session_state.data["kronologi_detail"] = st.text_area("Detail Kronologi & Keadaan Korban", st.session_state.data["kronologi_detail"])

    c1, c2 = st.columns(2)
    st.session_state.data["akibat_klien"] = c1.text_input("Akibat bagi Klien", st.session_state.data["akibat_klien"])
    st.session_state.data["akibat_masyarakat"] = c2.text_input("Akibat bagi Masyarakat", st.session_state.data["akibat_masyarakat"])

    st.session_state.data["sikap_klien_bab7"] = st.text_area("Sikap dan Tanggapan Klien", st.session_state.data["sikap_klien_bab7"])

    st.markdown("---")
    st.info("🤖 **AI Assistant: Analisis Hukum & Latar Belakang**")
    
    if st.button("Generate Analisis Hukum (Bab VI)"):
        st.session_state.data["ai_analisis_hukum"] = generate_narasi_ai("hukum", st.session_state.data)

    st.session_state.data["ai_analisis_hukum"] = st.text_area(
        "Hasil Analisis Otomatis (Dapat diedit sebelum dicetak):",
        st.session_state.data["ai_analisis_hukum"],
        height=150
    )

elif halaman == "2. Identitas Klien":
    st.header("👤 Identitas Lengkap Klien")

    st.markdown("---")

    c1, c2, c3 = st.columns(3)
    st.session_state.data["nama_klien"] = c1.text_input(
        "Nama Klien",
        value=st.session_state.data["nama_klien"],
        key="nama_klien_input"
    )
    st.session_state.data["no_reg_bin"] = c2.text_input(
        "No. Reg. Bin",
        value=st.session_state.data["no_reg_bin"],
        key="no_reg_bin_input"
    )
    st.session_state.data["ttl_klien"] = c3.text_input(
        "Tempat, Tgl Lahir Klien",
        value=st.session_state.data["ttl_klien"],
        key="ttl_klien_input"
    )

    c1, c2, c3 = st.columns(3)
    idx_jk = 0 if st.session_state.data["jk_klien"] == "Laki-laki" else 1
    st.session_state.data["jk_klien"] = c1.selectbox(
        "Jenis Kelamin",
        ["Laki-laki", "Perempuan"],
        index=idx_jk,
        key="jk_klien_select"
    )
    st.session_state.data["suku"] = c2.text_input(
        "Suku",
        value=st.session_state.data["suku"],
        key="suku_input"
    )
    st.session_state.data["warga_negara"] = c3.text_input(
        "Warga Negara",
        value=st.session_state.data["warga_negara"],
        key="warga_negara_input"
    )

    c1, c2, c3 = st.columns(3)
    st.session_state.data["agama"] = c1.text_input(
        "Agama",
        value=st.session_state.data["agama"],
        key="agama_input"
    )
    st.session_state.data["pend_terakhir"] = c2.text_input(
        "Pendidikan Terakhir",
        value=st.session_state.data["pend_terakhir"],
        key="pend_terakhir_input"
    )
    st.session_state.data["kerja_terakhir"] = c3.text_input(
        "Pekerjaan Terakhir",
        value=st.session_state.data["kerja_terakhir"],
        key="kerja_terakhir_input"
    )

    st.session_state.data["alamat_klien"] = st.text_area(
        "Alamat Lengkap Klien",
        value=st.session_state.data["alamat_klien"],
        key="alamat_klien_area"
    )

    c1, c2 = st.columns(2)
    st.session_state.data["ciri_khusus"] = c1.text_input(
        "Ciri-ciri Khusus",
        value=st.session_state.data["ciri_khusus"],
        key="ciri_khusus_input"
    )
    st.session_state.data["status_nikah"] = c2.text_input(
        "Status Pernikahan",
        value=st.session_state.data["status_nikah"],
        key="status_nikah_input"
    )

    st.session_state.data["riwayat_nikah_klien"] = st.text_area(
        "Riwayat Pernikahan Klien",
        value=st.session_state.data["riwayat_nikah_klien"],
        key="riwayat_nikah_klien_area"
    )

    st.subheader("Riwayat Pendidikan & Konsumsi")
    c1, c2, c3, c4 = st.columns(4)
    st.session_state.data["sd"] = c1.text_input(
        "Sekolah Dasar",
        value=st.session_state.data["sd"],
        key="sd_input"
    )
    st.session_state.data["smp"] = c2.text_input(
        "SMP",
        value=st.session_state.data["smp"],
        key="smp_input"
    )
    st.session_state.data["sma"] = c3.text_input(
        "SMA/SMK",
        value=st.session_state.data["sma"],
        key="sma_input"
    )
    st.session_state.data["pt"] = c4.text_input(
        "Perguruan Tinggi",
        value=st.session_state.data["pt"],
        key="pt_input"
    )

    c1, c2, c3 = st.columns(3)
    st.session_state.data["riwayat_rokok"] = c1.text_input(
        "Riwayat Rokok",
        value=st.session_state.data["riwayat_rokok"],
        key="riwayat_rokok_input"
    )
    st.session_state.data["riwayat_alkohol"] = c2.text_input(
        "Riwayat Alkohol",
        value=st.session_state.data["riwayat_alkohol"],
        key="riwayat_alkohol_input"
    )
    st.session_state.data["riwayat_napza"] = c3.text_input(
        "Riwayat NAPZA",
        value=st.session_state.data["riwayat_napza"],
        key="riwayat_napza_input"
    )

    # Opsional: Tambah AI helper untuk halaman ini (mirip halaman 1)
    st.markdown("---")
    st.info("🤖 **AI Assistant: Analisis Identitas Klien**")
    if st.button("Generate Analisis Klien"):
        st.session_state.data["ai_analisis_klien"] = generate_narasi_ai("klien", st.session_state.data)

    if "ai_analisis_klien" in st.session_state.data:
        st.session_state.data["ai_analisis_klien"] = st.text_area(
            "Hasil Analisis Otomatis (Dapat diedit):",
            value=st.session_state.data["ai_analisis_klien"],
            height=150,
            key="ai_analisis_klien_area"
        )

    st.markdown("---")
    st.info("🤖 **AI Assistant: Profil & Riwayat Klien**")
    if st.button("Generate Analisis Profil Klien"):
        st.session_state.data["ai_analisis_klien"] = generate_narasi_ai(
            "klien", st.session_state.data
        )

    st.session_state.data["ai_analisis_klien"] = st.text_area(
        "Hasil Analisis Otomatis:",
        st.session_state.data["ai_analisis_klien"],
        height=150,
        key="ai_analisis_klien_output"
        )
    
# Halaman 3
elif halaman == "3. Keluarga (Ayah & Ibu)":
    st.header("👨‍👩‍👦 Data Orang Tua")

    with st.expander("Data Ayah Kandung", expanded=True):
        c1, c2, c3 = st.columns(3)
        st.session_state.data["nama_ayah"] = c1.text_input(
            "Nama Ayah", st.session_state.data["nama_ayah"]
        )
        st.session_state.data["ttl_ayah"] = c2.text_input(
            "TTL Ayah", st.session_state.data["ttl_ayah"]
        )
        st.session_state.data["agama_ayah"] = c3.text_input(
            "Agama Ayah", st.session_state.data["agama_ayah"]
        )

        c1, c2, c3 = st.columns(3)
        st.session_state.data["suku_ayah"] = c1.text_input(
            "Suku Ayah", st.session_state.data["suku_ayah"]
        )
        st.session_state.data["pend_ayah"] = c2.text_input(
            "Pendidikan Ayah", st.session_state.data["pend_ayah"]
        )
        st.session_state.data["kerja_ayah"] = c3.text_input(
            "Pekerjaan Ayah", st.session_state.data["kerja_ayah"]
        )

        c1, c2, c3 = st.columns(3)
        st.session_state.data["alamat_ayah"] = c1.text_input(
            "Alamat Ayah", st.session_state.data["alamat_ayah"]
        )
        st.session_state.data["no_hp_ayah"] = c2.text_input(
            "No HP Ayah", st.session_state.data["no_hp_ayah"]
        )
        st.session_state.data["ket_ayah"] = c3.text_input(
            "Keterangan Ayah (Hidup/Mati)", st.session_state.data["ket_ayah"]
        )

    with st.expander("Data Ibu Kandung", expanded=True):
        c1, c2, c3 = st.columns(3)
        st.session_state.data["nama_ibu"] = c1.text_input(
            "Nama Ibu", st.session_state.data["nama_ibu"]
        )
        st.session_state.data["ttl_ibu"] = c2.text_input(
            "TTL Ibu", st.session_state.data["ttl_ibu"]
        )
        st.session_state.data["agama_ibu"] = c3.text_input(
            "Agama Ibu", st.session_state.data["agama_ibu"]
        )

        c1, c2, c3 = st.columns(3)
        st.session_state.data["suku_ibu"] = c1.text_input(
            "Suku Ibu", st.session_state.data["suku_ibu"]
        )
        st.session_state.data["pend_ibu"] = c2.text_input(
            "Pendidikan Ibu", st.session_state.data["pend_ibu"]
        )
        st.session_state.data["kerja_ibu"] = c3.text_input(
            "Pekerjaan Ibu", st.session_state.data["kerja_ibu"]
        )

        c1, c2, c3 = st.columns(3)
        st.session_state.data["alamat_ibu"] = c1.text_input(
            "Alamat Ibu", st.session_state.data["alamat_ibu"]
        )
        st.session_state.data["no_hp_ibu"] = c2.text_input(
            "No HP Ibu", st.session_state.data["no_hp_ibu"]
        )
        st.session_state.data["ket_ibu"] = c3.text_input(
            "Keterangan Ibu (Hidup/Mati)", st.session_state.data["ket_ibu"]
        )
# Halaman 4
elif halaman == "4. Penjamin & Saudara":
    st.header("🤝 Penjamin & Saudara Kandung")

    opsi = st.radio(
        "Penjamin adalah:", ["Ayah", "Ibu", "Orang Lain"], horizontal=True
    )

    st.markdown("---")
    if opsi == "Orang Lain":
        st.subheader("Data Penjamin (Pihak Lain)")
        c1, c2, c3 = st.columns(3)
        st.session_state.data["nama_pjm"] = c1.text_input(
            "Nama Penjamin", st.session_state.data["nama_pjm"]
        )
        st.session_state.data["ttl_pjm"] = c2.text_input(
            "TTL Penjamin", st.session_state.data["ttl_pjm"]
        )
        st.session_state.data["agama_pjm"] = c3.text_input(
            "Agama Penjamin", st.session_state.data["agama_pjm"]
        )

        c1, c2, c3 = st.columns(3)
        st.session_state.data["suku_pjm"] = c1.text_input(
            "Suku Penjamin", st.session_state.data["suku_pjm"]
        )
        st.session_state.data["pend_pjm"] = c2.text_input(
            "Pendidikan Penjamin", st.session_state.data["pend_pjm"]
        )
        st.session_state.data["kerja_pjm"] = c3.text_input(
            "Pekerjaan Penjamin", st.session_state.data["kerja_pjm"]
        )

        c1, c2, c3 = st.columns(3)
        st.session_state.data["alamat_pjm"] = c1.text_input(
            "Alamat Penjamin", st.session_state.data["alamat_pjm"]
        )
        st.session_state.data["hub_pjm"] = c2.text_input(
            "Hubungan Penjamin", st.session_state.data["hub_pjm"]
        )
        st.session_state.data["no_hp_pjm"] = c3.text_input(
            "No HP Penjamin", st.session_state.data["no_hp_pjm"]
        )

        c1, c2 = st.columns(2)
        st.session_state.data["layak_pjm"] = c1.text_area(
            "Kesiapan/Kelayakan Penjamin", st.session_state.data["layak_pjm"]
        )
        st.session_state.data["riwayat_nikah_pjm"] = c2.text_area(
            "Riwayat Perkawinan Penjamin",
            st.session_state.data["riwayat_nikah_pjm"],
        )
    else:
        st.success(
            f"✅ Data Penjamin akan diambil dari data **{opsi}** saat dicetak."
        )
        st.info(
            "Silakan isi 'Kesiapan Penjamin' dan 'Riwayat Perkawinan' di bawah ini:"
        )

        c1, c2 = st.columns(2)
        st.session_state.data["layak_pjm"] = c1.text_area(
            f"Kesiapan {opsi} sebagai Penjamin", st.session_state.data["layak_pjm"]
        )
        st.session_state.data["riwayat_nikah_pjm"] = c2.text_area(
            f"Riwayat Perkawinan {opsi} (Bab IV A)",
            st.session_state.data["riwayat_nikah_pjm"],
        )

    st.markdown("### Data Saudara Kandung")
    st.caption(
        "Tambah/hapus baris dengan tombol di bawah tabel. Isi kolom sesuai kebutuhan."
    )

    if (
        "saudara_data" not in st.session_state.data
        or len(st.session_state.data["saudara_data"].columns) != 7
    ):
        default_data = [
            {
                "No": 1,
                "Nama": "",
                "Usia": "",
                "JK": "L",
                "Pendidikan": "",
                "Pekerjaan": "",
                "Keterangan": "",
            },
            {
                "No": 2,
                "Nama": "",
                "Usia": "",
                "JK": "L",
                "Pendidikan": "",
                "Pekerjaan": "",
                "Keterangan": "",
            },
            {
                "No": 3,
                "Nama": "",
                "Usia": "",
                "JK": "L",
                "Pendidikan": "",
                "Pekerjaan": "",
                "Keterangan": "",
            },
        ]
        st.session_state.data["saudara_data"] = pd.DataFrame(default_data)

    edited_df = st.data_editor(
        st.session_state.data["saudara_data"],
        num_rows="dynamic",
        column_config={
            "No": st.column_config.NumberColumn(
                "No", min_value=1, step=1, required=True, format="%d"
            ),
            "Nama": st.column_config.TextColumn("Nama Lengkap", required=True),
            "Usia": st.column_config.NumberColumn(
                "Usia (tahun)", min_value=0, step=1, format="%d"
            ),
            "JK": st.column_config.SelectboxColumn(
                "Jenis Kelamin", options=["L", "P"], required=True
            ),
            "Pendidikan": st.column_config.TextColumn("Pendidikan Terakhir"),
            "Pekerjaan": st.column_config.TextColumn("Pekerjaan Saat Ini"),
            "Keterangan": st.column_config.TextColumn("Keterangan Tambahan"),
        },
        hide_index=False,
        use_container_width=True,
        key="saudara_editor",
    )

    st.session_state.data["saudara_data"] = edited_df

    st.markdown("---")
    st.info("🤖 **AI Assistant: Analisis Kelayakan Penjamin**")
    if st.button("Generate Analisis Penjamin"):
        st.session_state.data["ai_analisis_penjamin"] = generate_narasi_ai(
            "penjamin", st.session_state.data
        )

    st.session_state.data["ai_analisis_penjamin"] = st.text_area(
        "Hasil Analisis Otomatis:",
        st.session_state.data["ai_analisis_penjamin"],
        height=150,
    )
# Halaman 5
elif halaman == "5. Riwayat & Lingkungan":
    st.header("📈 Riwayat Pembinaan & Lingkungan")

    c1, c2 = st.columns(2)
    st.session_state.data["bina_pribadi"] = c1.text_area(
        "Pembinaan Kepribadian", st.session_state.data["bina_pribadi"]
    )
    st.session_state.data["bina_mandiri"] = c2.text_area(
        "Pembinaan Kemandirian", st.session_state.data["bina_mandiri"]
    )

    c1, c2 = st.columns(2)
    st.session_state.data["minat_bakat"] = c1.text_area(
        "Minat & Bakat", st.session_state.data["minat_bakat"]
    )
    st.session_state.data["relasi_masyarakat"] = c2.text_area(
        "Relasi Sosial Masyarakat", st.session_state.data["relasi_masyarakat"]
    )

    st.session_state.data["kondisi_alam"] = st.text_area(
        "Kondisi Alam/Lingkungan", st.session_state.data["kondisi_alam"]
    )

    st.markdown("---")
    st.info("🤖 **AI Assistant: Analisis Perkembangan & Lingkungan**")
    if st.button("Generate Analisis Lingkungan"):
        st.session_state.data["ai_analisis_lingkungan"] = generate_narasi_ai(
            "lingkungan", st.session_state.data
        )

    st.session_state.data["ai_analisis_lingkungan"] = st.text_area(
        "Hasil Analisis Otomatis:",
        st.session_state.data["ai_analisis_lingkungan"],
        height=150,
    )

# Halaman 6
elif halaman == "6. Asesmen (RRI)":
    st.header("📊 Asesmen Risiko Residivisme Indonesia (RRI)")

    with st.form(key="form_rri", clear_on_submit=False):

        r1 = st.radio(
            "1. Usia saat pertama kali melakukan tindak pidana",
            options=["0: > 21 thn", "1: 18–21 thn", "2: < 18 thn"],
            key="r1"
        )
    
        r2 = st.radio(
            "2. Apakah ada catatan kriminal sebelumnya?",
            options=["0: Tidak", "1: Ada"],
            key="r2"
        )
        
        r3 = st.radio(
            "3. Berapa jumlah pasal yang dikenakan?",
            options=["0: 1-2 pasal", "1: 3 pasal atau lebih"],
            key="r3"
        )
        
        r4 = st.radio(
            "4. Pernahkah mengalami kesulitan atau melanggar aturan selama program reintegrasi?",
            options=["0: Tidak", "1: Ya"],
            key="r4"
        )
        
        r5 = st.radio(
            "5. Sebelum kasus sekarang, berapa kali Anda pernah mendapat vonis pidana (tidak hitung pasal)?",
            options=["0: Tidak pernah", "1: 1-2 kali", "2: 3 kali atau lebih"],
            key="r5"
        )
        
        r6 = st.radio(
            "6. Selama di Rutan/Lapas/Bapas, pernahkah ada catatan perilaku kurang baik?",
            options=["0: Tidak", "1: Ya"],
            key="r6"
        )
        
        r7 = st.radio(
            "7. Apakah Anda pernah mengalami gangguan mental atau kecemasan?",
            options=["0: Tidak", "1: Ya"],
            key="r7"
        )
        
        r8 = st.radio(
            "8. Apakah ada keluarga dekat atau pasangan yang pernah berurusan dengan hukum?",
            options=["0: Tidak", "1: Ya"],
            key="r8"
        )
        
        r9 = st.radio(
            "9. Pernahkah menggunakan alkohol berlebihan atau narkotika? (pilih yang paling sesuai)",
            options=["0: Tidak pernah", "1: Hanya alkohol", "2: Hanya narkotika", "3: Keduanya"],
            key="r9"    
        )
        
        r10 = st.radio(
            "10. Pernahkah menganggur terus-menerus selama 12 bulan atau lebih?",
            options=["0: Tidak", "1: Ya"],
            key="r10"
        )
        
        col1, col2 = st.columns(2)
        with col1:
            gender = st.selectbox("Jenis Kelamin", ["Laki-laki", "Perempuan"], key="gender_select")
        with col2:
            kasus_narkotika = st.checkbox("Kasus terkait narkotika/obat terlarang?", key="kasus_narkotika_checkbox")
                     
        st.subheader("Bagian B: Faktor Tambahan")
        centang_b = 0
        if st.checkbox("Tindak pidana sekarang terasa lebih berat dibanding sebelumnya?"): centang_b += 1
        if st.checkbox("Sebelum usia 15 tahun, pernah melakukan tindak kekerasan?"): centang_b += 1
        if st.checkbox("Pernah mengalami/menjadi bagian kekerasan dalam rumah tangga?"): centang_b += 1
        if st.checkbox("Kasus melibatkan kekerasan, seksual, terorisme, atau separatisme?"): centang_b += 1

        centang_c = 0
        if gender == "Perempuan":
            st.subheader("Bagian C: Khusus Perempuan")
            if st.checkbox("Pernah melahirkan sebelum usia 20 tahun?"): centang_c += 1
            if st.checkbox("Pernah kesulitan merawat/membesarkan anak?"): centang_c += 1
            if st.checkbox("Pernah terlibat dunia prostitusi?"): centang_c += 1

        centang_d = 0
        if kasus_narkotika:
            st.subheader("Bagian D: Khusus Kasus Narkotika")      
            if st.checkbox("Pernah terlibat jaringan pengedar narkotika?"): centang_d += 1
            if st.checkbox("Merasa penggunaan narkoba oleh diri sendiri wajar/bisa dibenarkan?"): centang_d += 1
            if st.checkbox("Ada riwayat penggunaan narkoba rutin/berulang?"): centang_d += 1
            
        # TOMBOL SUBMIT HARUS DI DALAM WITH FORM
        submit_button = st.form_submit_button(
            "Hitung Skor & Tampilkan Hasil",
            type="primary",
            use_container_width=True
        )  # ← tutup dengan ) di sini

    # ────────────────────────────────────────────────
    # Proses hasil (di luar form, indent kembali ke level elif halaman)
    if submit_button:
        
        def get_score(val):
            if not val:
                return 0
            try:
                return int(val.split(":", 1)[0].strip())
            except:
                return 0

        skor_rri = 0
        skor_rri += get_score(r1)
        skor_rri += get_score(r2)
        skor_rri += get_score(r3)
        skor_rri += get_score(r4)
        skor_rri += get_score(r5)
        skor_rri += get_score(r6)
        skor_rri += get_score(r7)
        skor_rri += get_score(r8)
        skor_rri += get_score(r9)
        skor_rri += get_score(r10)

        skor_total = skor_rri + centang_b + centang_c + centang_d

        if skor_total <= 5:
            kategori = "RENDAH"
        elif skor_total <= 10:
            kategori = "SEDANG"
        else:
            kategori = "TINGGI"

        st.session_state.data["skor_rri"] = skor_total
        st.session_state.data["kat_rri"] = kategori

        st.metric("Total Skor RRI", skor_total, delta=kategori)

        st.write(f"**Skor dari 10 pertanyaan utama:** {skor_rri}")
        st.write(f"**Kontribusi Bagian B:** +{centang_b}")
        if gender == "Perempuan":
            st.write(f"**Kontribusi Bagian C:** +{centang_c}")
        if kasus_narkotika:
            st.write(f"**Kontribusi Bagian D:** +{centang_d}")

        st.success(f"Penilaian selesai! Skor total: {skor_total}")
            
# Halaman 6.1
elif halaman == "6. Asesmen Faktor Kriminogenik":
    st.header("Asesmen Faktor Kriminogenik (Kebutuhan Pembinaan)")

    st.info("Isi 29 pertanyaan berikut sesuai kondisi klien. Skor akan dihitung otomatis setelah submit.")

    with st.form(key="form_kriminogenik"):

        # Buat dict responses untuk 1-29
        responses = {}

        # Domain 1: Keluarga & Pernikahan (1-3)
        st.markdown("**Keluarga dan Pernikahan**")
        responses[1] = st.radio("Bagaimana hubungan Klien dengan pasangan saat ini? Apakah saling mendukung dan harmonis?", ["Ya, baik dan saling mendukung (0)", "Sebagian baik, ada masalah (1)", "Tidak baik, sering konflik (2)"], horizontal=True, index=0)
        responses[1] = 0 if "Ya" in responses[1] else (1 if "Sebagian" in responses[1] else 2)

        responses[2] = st.radio("Bagaimana hubungan Klien dengan orang tua atau wali?", ["Ya, baik dan mendukung (0)", "Sebagian baik (1)", "Tidak baik/konflik (2)"], horizontal=True, index=0)
        responses[2] = 0 if "Ya" in responses[2] else (1 if "Sebagian" in responses[2] else 2)

        responses[3] = st.radio("Bagaimana hubungan dengan anggota keluarga lainnya?", ["Ya, baik dan harmonis (0)", "Sebagian baik (1)", "Tidak baik/konflik (2)"], horizontal=True, index=0)
        responses[3] = 0 if "Ya" in responses[3] else (1 if "Sebagian" in responses[3] else 2)

        # Domain 2: Pendidikan & Pekerjaan (4-11)
        st.markdown("**Pendidikan dan Pekerjaan**")
        responses[4] = st.radio("Apakah Klien bisa membaca dan menulis dengan lancar?", ["Ya (0)", "Tidak / kesulitan (1)"], horizontal=True, index=0)
        responses[4] = 0 if "Ya" in responses[4] else 1

        responses[5] = st.radio("Apakah Klien menyelesaikan pendidikan minimal SMA/sederajat atau lebih tinggi?", ["Ya (0)", "Tidak / belum selesai (1)"], horizontal=True, index=0)
        responses[5] = 0 if "Ya" in responses[5] else 1

        responses[6] = st.radio("Apakah Klien menganggur sebelum kasus pidana sekarang?", ["Tidak / bekerja atau aktif (0)", "Ya, menganggur (1)"], horizontal=True, index=0)
        responses[6] = 1 if "Ya" in responses[6] else 0

        responses[7] = st.radio("Apakah Klien mengikuti program pembinaan kemandirian di Lapas/Bapas dalam 12 bulan terakhir?", ["Ya, ikut aktif (0)", "Tidak / belum ikut (1)"], horizontal=True, index=0)
        responses[7] = 0 if "Ya" in responses[7] else 1

        responses[8] = st.radio("Apakah Klien sering menganggur atau tidak bekerja sekitar setengah waktu saat di masyarakat?", ["Tidak / jarang (0)", "Ya, sering (1)"], horizontal=True, index=0)
        responses[8] = 1 if "Ya" in responses[8] else 0

        responses[9] = st.radio("Apakah kegiatan di tempat kerja/sekolah terasa bermakna dan bermanfaat bagi Klien?", ["Ya, sangat bermakna (0)", "Perlu dikembangkan (1)", "Tidak bermakna (2)"], horizontal=True, index=0)
        responses[9] = 0 if "Ya" in responses[9] else (1 if "Perlu" in responses[9] else 2)

        responses[10] = st.radio("Bagaimana hubungan Klien dengan rekan kerja/teman sekolah/kuliah?", ["Ya, baik dan harmonis (0)", "Perlu dikembangkan (1)", "Tidak baik/konflik (2)"], horizontal=True, index=0)
        responses[10] = 0 if "Ya" in responses[10] else (1 if "Perlu" in responses[10] else 2)

        responses[11] = st.radio("Bagaimana hubungan Klien dengan atasan di tempat kerja atau pengajar di sekolah?", ["Ya, baik dan saling menghargai (0)", "Perlu dikembangkan (1)", "Tidak baik/konflik (2)"], horizontal=True, index=0)
        responses[11] = 0 if "Ya" in responses[11] else (1 if "Perlu" in responses[11] else 2)

        st.markdown("**Penggunaan Narkotika, Obat Terlarang & Alkohol**")
        responses[12] = st.radio("Sebelum masalah hukum ini, apakah Klien secara rutin menggunakan narkotika/obat terlarang atau alkohol?", ["Tidak (0)", "Ya (1)"], horizontal=True, index=0)
        responses[12] = 1 if "Ya" in responses[12] else 0

        responses[13] = st.radio("Apakah tindak pidana yang dilakukan terkait atau dipengaruhi oleh penggunaan narkotika/alkohol?", ["Tidak (0)", "Ya (1)"], horizontal=True, index=0)
        responses[13] = 1 if "Ya" in responses[13] else 0

        responses[14] = st.radio("Apakah Klien pernah menggunakan narkotika/alkohol selama di Lapas/Rutan atau program reintegrasi?", ["Tidak (0)", "Ya (1)"], horizontal=True, index=0)
        responses[14] = 1 if "Ya" in responses[14] else 0

        responses[15] = st.radio("Apakah penggunaan tersebut berdampak negatif pada pekerjaan atau pendidikan Klien?", ["Tidak (0)", "Ya (1)"], horizontal=True, index=0)
        responses[15] = 1 if "Ya" in responses[15] else 0

        responses[16] = st.radio("Apakah penggunaan tersebut memengaruhi hubungan dengan pasangan atau keluarga?", ["Tidak (0)", "Ya (1)"], horizontal=True, index=0)
        responses[16] = 1 if "Ya" in responses[16] else 0

        responses[17] = st.radio("Apakah penggunaan tersebut berdampak negatif pada kesehatan Klien?", ["Tidak (0)", "Ya (1)"], horizontal=True, index=0)
        responses[17] = 1 if "Ya" in responses[17] else 0

        st.markdown("**Hubungan Sosial**")
        responses[18] = st.radio("Apakah tindak pidana sekarang dilakukan bersama teman atau rekan?", ["Tidak (0)", "Ya (1)"], horizontal=True, index=0)
        responses[18] = 1 if "Ya" in responses[18] else 0

        responses[19] = st.radio("Apakah Klien anggota kelompok/organisasi yang melakukan aktivitas kriminal?", ["Tidak (0)", "Ya (1)"], horizontal=True, index=0)
        responses[19] = 1 if "Ya" in responses[19] else 0

        responses[20] = st.radio("Apakah Klien memiliki teman dan rekan yang pro-sosial (positif dan mendukung hukum)?", ["Ya, banyak (0)", "Terbatas (1)", "Tidak / kebanyakan negatif (2)"], horizontal=True, index=0)
        responses[20] = 0 if "Ya" in responses[20] else (1 if "Terbatas" in responses[20] else 2)

        st.markdown("**Waktu Luang dan Rekreasi**")
        responses[21] = st.radio("Apakah Klien sering terlibat dalam kegiatan yang konstruktif dan bermanfaat?", ["Ya (0)", "Tidak / jarang (1)"], horizontal=True, index=0)
        responses[21] = 0 if "Ya" in responses[21] else 1

        responses[22] = st.radio("Apakah Klien memiliki terlalu banyak waktu luang yang tidak terisi positif?", ["Tidak (0)", "Ya (1)"], horizontal=True, index=0)
        responses[22] = 1 if "Ya" in responses[22] else 0

        st.markdown("**Manajemen Keuangan**")
        responses[23] = st.radio("Apakah kesulitan keuangan menjadi salah satu pendorong tindak pidana sekarang?", ["Tidak (0)", "Ya (1)"], horizontal=True, index=0)
        responses[23] = 1 if "Ya" in responses[23] else 0

        responses[24] = st.radio("Apakah Klien memiliki utang yang sulit dibayar?", ["Tidak (0)", "Ya (1)"], horizontal=True, index=0)
        responses[24] = 1 if "Ya" in responses[24] else 0

        st.markdown("**Sikap Anti-Sosial / Pandangan Terhadap Tindak Kriminal**")
        responses[25] = st.radio("Apakah Klien memiliki pandangan negatif terhadap sistem pidana/peradilan?", ["Tidak (0)", "Ya (1)"], horizontal=True, index=0)
        responses[25] = 1 if "Ya" in responses[25] else 0

        responses[26] = st.radio("Apakah Klien bisa merasakan empati atau menyesal terhadap dampak pada korban?", ["Ya, punya empati (0)", "Sedikit (1)", "Tidak / tidak merasa (2)"], horizontal=True, index=0)
        responses[26] = 0 if "Ya" in responses[26] else (1 if "Sedikit" in responses[26] else 2)

        responses[27] = st.radio("Apakah Klien memiliki riwayat kejahatan kekerasan atau kekerasan seksual yang berulang?", ["Tidak (0)", "Perlu perhatian / ada riwayat (1)", "Ya, berulang (2)"], horizontal=True, index=0)
        responses[27] = 0 if "Tidak" in responses[27] else (1 if "Perlu" in responses[27] else 2)

        responses[28] = st.radio("Apakah Klien memiliki sikap negatif terhadap rehabilitasi atau program layanan?", ["Tidak / terbuka (0)", "Ya / menolak (1)"], horizontal=True, index=0)
        responses[28] = 1 if "Ya" in responses[28] else 0

        responses[29] = st.radio("Apakah Klien meyakini kejahatan adalah cara sah untuk memenuhi kebutuhan?", ["Tidak (0)", "Ya (1)"], horizontal=True, index=0)
        responses[29] = 1 if "Ya" in responses[29] else 0

        st.subheader("Bagian B.1 – Pertimbangan Tindak Pidana Tertentu (centang jika sesuai)")
        centang_b1 = []
        if st.checkbox("Memiliki sejarah melakukan kekerasan"): centang_b1.append(1)
        if st.checkbox("Kejahatan karena tindakan tidak bermoral"): centang_b1.append(2)
        if st.checkbox("Kejahatan karena kekerasan ekstremis atau terorisme"): centang_b1.append(3)
        if st.checkbox("Pernah menjadi korban kekerasan (termasuk KDRT)"): centang_b1.append(4)
        if st.checkbox("Terlibat dalam kejahatan luar biasa yang diproses peradilan khusus"): centang_b1.append(5)
        if st.checkbox("Kejahatan karena korupsi atau penipuan"): centang_b1.append(6)
        if st.checkbox("Kesulitan mengelola emosi"): centang_b1.append(7)
        if st.checkbox("Kejahatan rasial (berdasarkan ras/agama)"): centang_b1.append(8)
        if st.checkbox("Kejahatan terkait narkotika (penyimpanan, panen, impor)"): centang_b1.append(9)

        st.subheader("Bagian B.2 – Faktor Kebutuhan Tambahan (centang jika sesuai)")
        centang_b2 = []
        if st.checkbox("Ada ancaman terhadap diri dari pihak ketiga"): centang_b2.append(1)
        if st.checkbox("Berpotensi menjadi tunawisma setelah bebas"): centang_b2.append(2)
        if st.checkbox("Masalah akomodasi/tempat tinggal (selain tunawisma)"): centang_b2.append(3)
        if st.checkbox("Pernah/sedang menjadi target operasi aparat penegak hukum"): centang_b2.append(4)
        if st.checkbox("Masalah kepatuhan terhadap otoritas"): centang_b2.append(5)
        if st.checkbox("Kemampuan bersosialisasi kurang baik"): centang_b2.append(6)
        if st.checkbox("Kesulitan belajar"): centang_b2.append(7)
        if st.checkbox("Penyandang disabilitas"): centang_b2.append(8)
        if st.checkbox("Pernah memiliki masalah keimigrasian"): centang_b2.append(9)

        submit_krim = st.form_submit_button("Hitung Hasil Kriminogenik", type="primary", use_container_width=True)

    if submit_krim:
        skor_domain, total = hitung_skor_domain(responses)
        kat_domain = kategori_domain(skor_domain)
        kat_total, warna_total = kategori_total(total)
        faktor_dominan, max_skor = get_faktor_dominan(skor_domain)
        kesimpulan, rekomendasi = get_kesimpulan_dan_rekomendasi(
            total, kat_total, skor_domain, faktor_dominan, centang_b1, centang_b2
        )

        # Tampilkan hasil
        st.metric("Total Skor Kriminogenik", total, delta=kat_total, delta_color="normal")

        st.subheader("Skor per Domain")
        for domain, skor in skor_domain.items():
            kat, warna = kat_domain[domain]
            st.write(f"**{domain.replace('_', ' ').title()}**: {skor} → {kat}")

        st.markdown("### Kesimpulan")
        st.markdown(kesimpulan)

        st.markdown("### Rekomendasi Pembinaan")
        st.markdown(rekomendasi)

        # Buat visualisasi hasil kriminogenik

        fig, ax = plt.subplots(figsize=(10, 6))
        domains = list(skor_domain.keys())
        scores = list(skor_domain.values())
        colors = [kat_domain[d][1] for d in domains]
        ax.barh(domains, scores, color=colors)
        ax.set_xlabel("Skor")
        ax.set_title("Skor Faktor Kriminogenik per Domain")
        plt.tight_layout()
        
        st.pyplot(fig)        # Simpan ke session_state untuk download
        st.session_state.data["skor_krim_total"] = total
        st.session_state.data["kat_krim"] = kat_total
        st.session_state.data["faktor_dominan_krim"] = faktor_dominan
        st.session_state.data["kesimpulan_krim"] = kesimpulan
        st.session_state.data["rekomendasi_krim"] = rekomendasi

        st.success("Asesmen kriminogenik selesai!")
    
# Halaman 7
elif halaman == "7. Download":
    st.header("🏁 Cetak Laporan")

    with st.form(key="form_download"):
        st.info("Pastikan semua data sudah terisi lengkap dan narasi AI sudah digenerate sebelum mencetak.")

        sync_pjm = st.checkbox("Gunakan Data Ayah/Ibu sebagai Penjamin?", key="sync_pjm_check")

        pilihan_pjm = st.radio(
            "Pilih sumber Penjamin:",
            ["Manual / Orang Lain", "Ayah", "Ibu"],
            index=0,
            key="pilihan_pjm_radio"
        )

        # Tombol submit HARUS di dalam form & disimpan ke variabel
        submit_download = st.form_submit_button(
            "Generate dan Download Dokumen",
            type="primary",
            use_container_width=True
        )

    # Proses hanya jalan setelah tombol ditekan (di luar form)
    if submit_download:
        if validate_data(st.session_state.data):
            try:
                st.success("Data valid! Sedang membuat dokumen...")

                # Load template
                doc = DocxTemplate("template.docx")  # pastikan file ada di folder yang sama

                # Copy context dari session_state.data
                context = st.session_state.data.copy()

                # Fallback narasi kosong
                narasi_keys = [
                    "ai_analisis_hukum", "ai_analisis_klien", "ai_analisis_penjamin",
                    "ai_analisis_lingkungan", "ai_analisis_rel_kel", "ai_analisis_kronologi",
                    "ai_analisis_akibat", "ai_analisis_sikap_bab7", "ai_analisis_hasil_rekomendasi",
                    "ai_analisis_kesimpulan"
                ]
                for key in narasi_keys:
                    if key not in context or not context.get(key):
                        context[key] = "[Narasi belum digenerate]"

                # Sync penjamin kalau dicentang
                if sync_pjm and pilihan_pjm != "Manual / Orang Lain":
                    context = sync_penjamin_data(context, pilihan_pjm)

                # Format tanggal jadi teks Indonesia
                date_fields = ["tgl_surat", "tgl_tpp", "tgl_1per3", "tgl_1per2", "tgl_2per3"]
                for key in date_fields:
                    if key in context and context[key]:
                        context[key] = format_tgl_tulis_indo(context[key])

                # Saudara data jadi list dict untuk Jinja
                if "saudara_data" in context:
                    context["saudara_data"] = context["saudara_data"].to_dict("records")

                # Render template
                doc.render(context)

                # Buffer untuk download
                bio = io.BytesIO()
                doc.save(bio)
                bio.seek(0)

                # Nama file dinamis
                nama_klien = context.get("nama_klien", "Klien").replace(" ", "_")
                tgl_now = datetime.now().strftime("%Y%m%d_%H%M")
                file_name = f"LITMAS_{nama_klien}_{tgl_now}.docx"

                # Tombol download langsung muncul setelah berhasil
                st.download_button(
                    label="📥 Unduh File LITMAS Sekarang",
                    data=bio,
                    file_name=file_name,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )

                st.success("Dokumen berhasil dibuat! Silakan klik tombol unduh di atas.")

            except FileNotFoundError:
                st.error("File 'template.docx' tidak ditemukan. Pastikan ada di folder yang sama dengan app.")
            except Exception as e:
                st.error(f"Terjadi kesalahan saat generate dokumen: {str(e)}")
                st.exception(e)  # tampilkan traceback lengkap untuk debug

        else:
            st.error("Data tidak lengkap! Lengkapi field wajib terlebih dahulu.")
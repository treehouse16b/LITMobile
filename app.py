import streamlit as st
from docxtpl import DocxTemplate
import io
import os
import pandas as pd
from datetime import datetime
import google.generativeai as genai

# --- 1. KONFIGURASI AI ---
# Masukkan API Key Gemini Anda di sini
API_KEY = "AIzaSyCDZQHto8k6dskFmh6QuV3F4WRwwSHJR68" 

def generate_full_litmas_narasi(data):
    try:
        genai.configure(api_key=API_KEY)
        model = genai.GenerativeModel('gemini-1.5-flash')
        
        prompt = f"""
        Anda adalah Pembimbing Kemasyarakatan (PK) Ahli. Ubah data input sederhana menjadi narasi paragraf formal kedinasan (Bahasa Indonesia resmi).
        
        DATA INPUT:
        - Nikah Penjamin: {data['riwayat_nikah_pjm']}
        - Relasi Keluarga: {data['relasi_keluarga']}
        - Relasi Masyarakat: {data['relasi_masyarakat']}
        - Kondisi Lingkungan (Bab V): {data['kondisi_lingkungan']}
        - Kronologi & Korban (Bab VI): {data['kronologi_detail']}, {data['korban_detail']}
        - Akibat TP (Bab VI): {data['akibat_tp']}
        - Sikap Klien (Bab VII): {data['sikap_klien_bab7']}
        - Asesmen & Pembinaan: RRI {data['hasil_rri']}, Kriminogenik {data['hasil_kriminogenik']}, Bina {data['bina_pribadi']} & {data['bina_mandiri']}
        - Program: {data['program']}

        TUGAS: Buat narasi paragraf untuk bagian berikut dengan penanda persis ###[TAG]###:
        ###N_NIKAH_PJM### : Narasi Bab IV A (Riwayat Perkawinan Penjamin).
        ###N_REL_KEL### : Narasi Bab IV B (Relasi dalam Keluarga).
        ###N_REL_MASY### : Narasi Bab IV C & V A (Relasi dalam Masyarakat).
        ###N_LINGKUNGAN### : Narasi Bab V B & C (Kondisi Alam & Keadaan Masyarakat).
        ###N_KRONOLOGI### : Narasi Bab VI A, B, C (Latar Belakang, Kronologis, Keadaan Korban).
        ###N_AKIBAT### : Narasi Bab VI D (Akibat terhadap Klien, Keluarga, Masyarakat).
        ###N_SIKAP_BAB7### : Narasi Bab VII (Sikap dan Tanggapan Klien).
        ###ANALISIS_PK### : Analisis Bab XI (Gunakan teori kriminologi RNR/Strain).
        ###KESIMPULAN### : Bab XII (Kelayakan program).
        ###REKOMENDASI### : Bab XIII (Saran konkret pengawasan).
        """
        response = model.generate_content(prompt)
        text = response.text
        
        def extract(tag):
            try: return text.split(f"###{tag}###")[1].split("###")[0].strip()
            except: return "Data tidak terproses, silakan isi manual."

        return {
            'n_nikah_pjm': extract('N_NIKAH_PJM'),
            'n_rel_kel': extract('N_REL_KEL'),
            'n_rel_masy': extract('N_REL_MASY'),
            'n_lingkungan': extract('N_LINGKUNGAN'),
            'n_kronologi': extract('N_KRONOLOGI'),
            'n_akibat': extract('N_AKIBAT'),
            'n_sikap_bab7': extract('N_SIKAP_BAB7'),
            'analisis_pk': extract('ANALISIS_PK'),
            'kesimpulan': extract('KESIMPULAN'),
            'rekomendasi_pk': extract('REKOMENDASI')
        }
    except Exception as e:
        st.error(f"Gagal AI: {e}")
        return {}

def format_tgl(tgl_obj):
    if not tgl_obj: return ""
    bulan_indo = {1: "Januari", 2: "Februari", 3: "Maret", 4: "April", 5: "Mei", 6: "Juni", 
                  7: "Juli", 8: "Agustus", 9: "September", 10: "Oktober", 11: "November", 12: "Desember"}
    return f"{tgl_obj.day} {bulan_indo[tgl_obj.month]} {tgl_obj.year}"

# --- 2. ANTARMUKA (UI) ---
st.set_page_config(page_title="Litmas Pro AI Final", layout="wide")
st.title("📂 Litmas Generator Terintegrasi AI (Versi Final 2026)")

with st.form("litmas_form"):
    # ADMINISTRASI
    st.header("🏢 I. Administrasi & Tahapan Pidana")
    c1, c2, c3 = st.columns(3)
    asal_upt = c1.text_input("Asal UPT")
    no_reg_litmas = c2.text_input("No. Reg LITMAS")
    no_surat = c3.text_input("No. Surat Permintaan")
    tgl_surat = c1.date_input("Tanggal Surat")
    perihal = c2.text_input("Perihal")
    program = c3.selectbox("Program Re-Integrasi", ["Pembebasan Bersyarat", "Cuti Bersyarat", "Cuti Menjelang Bebas", "Asimilasi"])
    
    c1, c2, c3, c4, c5 = st.columns(5)
    tgl_tpp = c1.date_input("Tgl Sidang TPP")
    tgl_1per3 = c2.date_input("Tgl 1/3")
    tgl_1per2 = c3.date_input("Tgl 1/2")
    tgl_2per3 = c4.date_input("Tgl 2/3")
    ekspirasi = c5.text_input("Tgl Ekspirasi (Bebas Murni)")

    # HUKUM
    st.header("⚖️ II. Data Hukum")
    c1, c2, c3 = st.columns(3)
    perkara = c1.text_input("Perkara")
    pasal = c2.text_input("Pasal")
    vonis = c3.text_input("Vonis")
    latarbelakang_tp = st.text_area("Latar Belakang TP (Input Sederhana)")
    kronologi_detail = st.text_area("Detail Kronologi & Keadaan Korban (Bab VI)")
    akibat_tp = st.text_area("Akibat bagi Klien/Masy (Bab VI)")

    # DATA KLIEN
    st.header("👤 III. Data Klien")
    c1, c2, c3 = st.columns(3)
    nama_klien = c1.text_input("Nama Klien")
    no_reg = c2.text_input("No. Reg. Bin")
    ttl_klien = c3.text_input("Tempat, Tgl Lahir Klien")
    jk_klien = c1.selectbox("Jenis Kelamin", ["Laki-laki", "Perempuan"])
    suku_klien = c2.text_input("Suku")
    warga_negara = c3.text_input("Warga Negara", "Indonesia")
    agama_klien = c1.text_input("Agama")
    pend_Klien = c2.text_input("Pendidikan Terakhir")
    kerja_klien = c3.text_input("Pekerjaan")
    alamat_klien = st.text_area("Alamat Lengkap Klien")
    ciri_klien = st.text_input("Ciri-ciri Khusus")
    status_nikah_klien = st.text_input("Status Pernikahan")
    riwayat_nikah = st.text_area("Riwayat Pernikahan Klien")
    
    c1, c2, c3, c4 = st.columns(4)
    sd = c1.text_input("Sekolah Dasar")
    smp = c2.text_input("SMP")
    sma = c3.text_input("SMA/SMK")
    perguruan_tinggi = c4.text_input("Perguruan Tinggi")

    c1, c2, c3, c4 = st.columns(4)
    riwayat_rokok = c1.text_input("Riwayat Rokok")
    riwayat_alkohol = c2.text_input("Riwayat Alkohol")
    riwayat_napza = c3.text_input("Riwayat NAPZA")
    minat_bakat = c4.text_input("Minat & Bakat")

    st.header("📈 IV. Pembinaan & Lingkungan")
    c1, c2 = st.columns(2)
    bina_pribadi = c1.text_area("Pembinaan Kepribadian")
    bina_mandiri = c2.text_area("Pembinaan Kemandirian")
    
    c1, c2 = st.columns(2)
    hasil_rri = c1.text_input("Hasil Skor RRI")
    hasil_kriminogenik = c2.text_input("Hasil Faktor Kriminogenik")
    
    c1, c2 = st.columns(2)
    relasi_keluarga = c1.text_area("Relasi dalam Keluarga (Bab IV B)")
    relasi_masyarakat = c2.text_area("Relasi dalam Masyarakat (Bab IV C)")
    kondisi_lingkungan = st.text_area("Kondisi Alam & Masyarakat (Bab V)")
    sikap_klien_bab7 = st.text_area("Sikap & Tanggapan Klien (Bab VII)")

    # KELUARGA & PENJAMIN
    st.header("👨‍👩‍👦 V. Keluarga & Penjamin")
    st.subheader("Saudara Kandung")
    df_saudara = pd.DataFrame([{"No": 1, "Nama": "", "JK": "L", "Usia": "", "Pendidikan": "", "Pekerjaan": "", "Keterangan": ""}])
    saudara_editor = st.data_editor(df_saudara, num_rows="dynamic")

    with st.expander("Detail Orang Tua"):
        c1, c2 = st.columns(2)
        with c1:
            nama_ayah = st.text_input("Nama Ayah")
            ttl_ayah = st.text_input("TTL Ayah")
            agama_ayah = st.text_input("Agama Ayah")
            suku_ayah = st.text_input("Suku/WN Ayah")
            pend_ayah = st.text_input("Pendidikan Ayah")
            kerja_ayah = st.text_input("Pekerjaan Ayah")
            alamat_ayah = st.text_area("Alamat Ayah")
            ket_ayah = st.text_input("Keterangan Ayah")
        with c2:
            nama_ibu = st.text_input("Nama Ibu")
            ttl_ibu = st.text_input("TTL Ibu")
            agama_ibu = st.text_input("Agama Ibu")
            suku_ibu = st.text_input("Suku/WN Ibu")
            pend_ibu = st.text_input("Pendidikan Ibu")
            kerja_ibu = st.text_input("Pekerjaan Ibu")
            alamat_ibu = st.text_area("Alamat Ibu")
            ket_ibu = st.text_input("Keterangan Ibu")

    with st.expander("Detail Penjamin"):
        c1, c2, c3 = st.columns(3)
        nama_pjm = c1.text_input("Nama Penjamin")
        ttl_pjm = c2.text_input("TTL Penjamin")
        agama_pjm = c3.text_input("Agama Penjamin")
        suku_pjm = c1.text_input("Suku/WN Penjamin")
        pend_pjm = c2.text_input("Pendidikan Penjamin")
        kerja_pjm = c3.text_input("Pekerjaan Penjamin")
        alamat_pjm = st.text_area("Alamat Penjamin")
        hub_pjm = c1.text_input("Hubungan Penjamin")
        no_hp_pjm = c2.text_input("No HP Penjamin")
        layak_pjm = st.text_area("Kesiapan Penjamin (Layak/Tidak)")
        riwayat_nikah_pjm = st.text_area("Riwayat Perkawinan Penjamin (Bab IV A)")

    submit = st.form_submit_button("🚀 GENERATE LITMAS FINAL")

# --- 3. PROSES & GENERATE ---
if submit:
    with st.spinner("AI sedang merangkai narasi profesional..."):
        narasi = generate_full_litmas_narasi({
            'riwayat_nikah_pjm': riwayat_nikah_pjm, 'relasi_keluarga': relasi_keluarga,
            'relasi_masyarakat': relasi_masyarakat, 'kondisi_lingkungan': kondisi_lingkungan,
            'kronologi_detail': kronologi_detail, 'korban_detail': "Data Korban: " + kronologi_detail,
            'akibat_tp': akibat_tp, 'sikap_klien_bab7': sikap_klien_bab7, 'nama_klien': nama_klien,
            'perkara': perkara, 'pasal': pasal, 'hasil_rri': hasil_rri, 'hasil_kriminogenik': hasil_kriminogenik,
            'bina_pribadi': bina_pribadi, 'bina_mandiri': bina_mandiri, 'program': program
        })
        
        list_saudara = []
        for _, row in saudara_editor.iterrows():
            list_saudara.append({'no': row['No'], 'nama': row['Nama'], 'jk': row['JK'], 'usia': row['Usia'], 'pend': row['Pendidikan'], 'kerja': row['Pekerjaan'], 'ket': row['Keterangan']})

        context = {
            # Administrasi
            'asal_upt': asal_upt, 'no_reg_litmas': no_reg_litmas, 'no_surat': no_surat, 'tgl_surat': format_tgl(tgl_surat),
            'perihal': perihal, 'program': program, 'tgl_tpp': format_tgl(tgl_tpp), 'tgl_1per3': format_tgl(tgl_1per3),
            'tgl_1per2': format_tgl(tgl_1per2), 'tgl_2per3': format_tgl(tgl_2per3), 'ekspirasi': ekspirasi, 'vonis': vonis,
            'tgl_buat': format_tgl(datetime.now()),
            # Klien
            'nama_klien': nama_klien, 'no_reg': no_reg, 'ttl_klien': ttl_klien, 'jk_klien': jk_klien, 'suku_klien': suku_klien,
            'warga_negara': warga_negara, 'agama_klien': agama_klien, 'pend_Klien': pend_Klien, 'sd': sd, 'smp': smp, 'sma': sma,
            'perguruan_tinggi': perguruan_tinggi, 'kerja_klien': kerja_klien, 'alamat_klien': alamat_klien, 'ciri_klien': ciri_klien,
            'status_nikah_klien': status_nikah_klien, 'riwayat_nikah': riwayat_nikah, 'riwayat_rokok': riwayat_rokok,
            'riwayat_alkohol': riwayat_alkohol, 'riwayat_napza': riwayat_napza, 'minat_bakat': minat_bakat,
            # Hukum & AI Narasi
            'perkara': perkara, 'pasal': pasal, 'latarbelakang_tp': latarbelakang_tp, 'hasil_rri': hasil_rri, 'hasil_kriminogenik': hasil_kriminogenik,
            'n_nikah_pjm': narasi.get('n_nikah_pjm'), 'n_rel_kel': narasi.get('n_rel_kel'), 'n_rel_masy': narasi.get('n_rel_masy'),
            'n_lingkungan': narasi.get('n_lingkungan'), 'n_kronologi': narasi.get('n_kronologi'), 'n_akibat': narasi.get('n_akibat'),
            'n_sikap_bab7': narasi.get('n_sikap_bab7'), 'analisis_pk': narasi.get('analisis_pk'), 'kesimpulan': narasi.get('kesimpulan'),
            'rekomendasi_pk': narasi.get('rekomendasi_pk'),
            # Ortu & Pjm
            'nama_ayah': nama_ayah, 'ttl_ayah': ttl_ayah, 'agama_ayah': agama_ayah, 'suku_ayah': suku_ayah, 'pend_ayah': pend_ayah, 'kerja_ayah': kerja_ayah, 'alamat_ayah': alamat_ayah, 'ket_ayah': ket_ayah,
            'nama_ibu': nama_ibu, 'ttl_ibu': ttl_ibu, 'agama_ibu': agama_ibu, 'suku_ibu': suku_ibu, 'pend_ibu': pend_ibu, 'kerja_ibu': kerja_ibu, 'alamat_ibu': alamat_ibu, 'ket_ibu': ket_ibu,
            'nama_pjm': nama_pjm, 'ttl_pjm': ttl_pjm, 'agama_pjm': agama_pjm, 'suku_pjm': suku_pjm, 'pend_pjm': pend_pjm, 'kerja_pjm': kerja_pjm, 'alamat_pjm': alamat_pjm, 'hub_pjm': hub_pjm, 'no_hp_pjm': no_hp_pjm, 'layak_pjm': layak_pjm,
            'saudara': list_saudara
        }

        try:
            doc = DocxTemplate("template.docx")
            doc.render(context)
            bio = io.BytesIO()
            doc.save(bio)
            bio.seek(0)
            st.success("✅ Litmas Berhasil Digenerate!")
            st.download_button("📥 DOWNLOAD FILE WORD", bio, f"Litmas_Final_{nama_klien}.docx")
        except Exception as e:
            st.error(f"Error Render Word: {e}")

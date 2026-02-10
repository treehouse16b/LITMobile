import streamlit as st
from docxtpl import DocxTemplate
import io
import os
import pandas as pd
from datetime import datetime
import google.generativeai as genai

# --- 1. KONFIGURASI AI ---
# Masukkan API Key Anda di sini
API_KEY = "AIzaSy..." 

def generate_analisis_ai(data):
    try:
        genai.configure(api_key=API_KEY)
        model = genai.GenerativeModel('gemini-1.5-flash')
        prompt = f"""
        Buatlah narasi analisis Penelitian Kemasyarakatan (Litmas) Bab XI yang mendalam.
        Data Klien: {data['nama_klien']}, Kasus: {data['perkara']} ({data['pasal']}).
        Latar Belakang & Kronologi: {data['latarbelakang_tp']} - {data['kronologi']}.
        Keadaan Korban: {data['keadaan_korban']}.
        Sikap Klien: {data['sikap_klien']}.
        Hasil Asesmen: RRI ({data['hasil_rri']}), Kriminogenik ({data['hasil_kriminogenik']}).
        Pembinaan: {data['bina_pribadi']} & {data['bina_mandiri']}.
        Program: {data['program']}.
        
        Tugas: Susun analisis profesional mengenai perubahan perilaku dan kelayakan program.
        """
        response = model.generate_content(prompt)
        return response.text
    except Exception as e:
        return f"Analisis otomatis gagal: {str(e)}. Silakan isi manual di Word."

def format_tgl(tgl_obj):
    if not tgl_obj: return ""
    bulan_indo = {1: "Januari", 2: "Februari", 3: "Maret", 4: "April", 5: "Mei", 6: "Juni", 
                  7: "Juli", 8: "Agustus", 9: "September", 10: "Oktober", 11: "November", 12: "Desember"}
    return f"{tgl_obj.day} {bulan_indo[tgl_obj.month]} {tgl_obj.year}"

# --- 2. TAMPILAN ANTARMUKA ---
st.set_page_config(page_title="Litmas Pro AI 2026", layout="wide")
st.title("📂 Generator Litmas Terintegrasi AI (Versi Lengkap)")

with st.form("litmas_form"):
    # --- ADMINISTRASI ---
    st.header("🏢 I. Administrasi & Tahapan Pidana")
    c1, c2, c3 = st.columns(3)
    asal_upt = c1.text_input("Asal UPT")
    no_reg_litmas = c2.text_input("No. Reg LITMAS")
    no_surat = c3.text_input("No. Surat Permintaan")
    tgl_surat = c1.date_input("Tanggal Surat Permintaan")
    perihal = c2.text_input("Perihal")
    program = c3.selectbox("Program", ["PB", "CB", "CMB", "Asimilasi"])
    
    c1, c2, c3, c4 = st.columns(4)
    tgl_tpp = c1.date_input("Tgl Sidang TPP")
    tgl_1per3 = c2.date_input("Tgl 1/3")
    tgl_1per2 = c3.date_input("Tgl 1/2")
    tgl_2per3 = c4.date_input("Tgl 2/3")

    # --- HUKUM ---
    st.header("⚖️ II. Data Hukum")
    c1, c2, c3 = st.columns(3)
    perkara = c1.text_input("Perkara")
    pasal = c2.text_input("Pasal")
    vonis = c3.text_input("Vonis (Lama Pidana)")
    ekspirasi = c1.text_input("Tanggal Ekspirasi (Bebas Murni)")
    latarbelakang_tp = st.text_area("Latar Belakang Tindak Pidana")
    kronologi = st.text_area("Kronologi Kejadian")
    keadaan_korban = st.text_area("Keadaan Korban")

    # --- DATA KLIEN ---
    st.header("👤 III. Data Klien")
    c1, c2, c3 = st.columns(3)
    nama_klien = c1.text_input("Nama Klien")
    no_reg = c2.text_input("No. Reg. Bin")
    ttl_klien = c3.text_input("Tempat, Tgl Lahir Klien")
    jk_klien = c1.selectbox("Jenis Kelamin", ["Laki-laki", "Perempuan"])
    suku_klien = c2.text_input("Suku Klien")
    warga_negara = c3.text_input("Warga Negara", "Indonesia")
    agama_klien = c1.text_input("Agama Klien")
    pend_klien = c2.text_input("Pendidikan Terakhir Klien")
    kerja_klien = c3.text_input("Pekerjaan Klien")
    alamat_klien = st.text_area("Alamat Klien")
    ciri_klien = st.text_input("Ciri-ciri Khusus Klien")
    
    st.subheader("🎓 Riwayat Pendidikan")
    c1, c2, c3, c4 = st.columns(4)
    sd = c1.text_input("SD")
    smp = c2.text_input("SMP")
    sma = c3.text_input("SMA")
    pt = c4.text_input("Perguruan Tinggi")

    st.subheader("💍 Status & Kebiasaan")
    c1, c2 = st.columns(2)
    status_nikah = c1.text_input("Status Pernikahan Klien")
    riwayat_nikah = c2.text_area("Riwayat Pernikahan")
    
    c1, c2, c3 = st.columns(3)
    r_rokok = c1.text_input("Riwayat Rokok")
    r_alkohol = c2.text_input("Riwayat Alkohol")
    r_napza = c3.text_input("Riwayat NAPZA")
    minat_bakat = st.text_input("Minat & Bakat")

    st.subheader("📈 Pembinaan & Asesmen")
    sikap_klien = st.text_area("Sikap & Tanggapan Klien")
    c1, c2 = st.columns(2)
    bina_pribadi = c1.text_area("Pembinaan Kepribadian")
    bina_mandiri = c2.text_area("Pembinaan Kemandirian")
    
    c1, c2 = st.columns(2)
    hasil_rri = c1.text_input("Hasil RRI")
    hasil_kriminogenik = c2.text_input("Hasil Kriminogenik")

    # --- KELUARGA & PENJAMIN ---
    st.header("👨‍👩‍👦 IV. Keluarga & Penjamin")
    
    st.subheader("👨‍👩‍👧‍👦 Tabel Susunan Keluarga (Saudara)")
    if "df_saudara" not in st.session_state:
        st.session_state.df_saudara = pd.DataFrame([{"No": 1, "Nama": "", "JK": "L", "Usia": "", "Pendidikan": "", "Pekerjaan": "", "Keterangan": ""}])
    saudara_editor = st.data_editor(st.session_state.df_saudara, num_rows="dynamic")

    with st.expander("Data Ayah Kandung"):
        c1, c2, c3 = st.columns(3)
        n_ayah = c1.text_input("Nama Ayah")
        ttl_ayah = c2.text_input("TTL Ayah")
        ag_ayah = c3.text_input("Agama Ayah")
        s_ayah = c1.text_input("Suku Ayah")
        p_ayah = c2.text_input("Pendidikan Ayah")
        k_ayah = c3.text_input("Pekerjaan Ayah")
        al_ayah = st.text_area("Alamat Ayah")
        ket_ayah = st.text_input("Keterangan Ayah")

    with st.expander("Data Ibu Kandung"):
        c1, c2, c3 = st.columns(3)
        n_ibu = c1.text_input("Nama Ibu")
        ttl_ibu = c2.text_input("TTL Ibu")
        ag_ibu = c3.text_input("Agama Ibu")
        s_ibu = c1.text_input("Suku Ibu")
        p_ibu = c2.text_input("Pendidikan Ibu")
        k_ibu = c3.text_input("Pekerjaan Ibu")
        al_ibu = st.text_area("Alamat Ibu")
        ket_ibu = st.text_input("Keterangan Ibu")

    with st.expander("Data Penjamin"):
        c1, c2, c3 = st.columns(3)
        n_pjm = c1.text_input("Nama Penjamin")
        ttl_pjm = c2.text_input("TTL Penjamin")
        ag_pjm = c3.text_input("Agama Penjamin")
        s_pjm = c1.text_input("Suku Penjamin")
        p_pjm = c2.text_input("Pendidikan Penjamin")
        k_pjm = c3.text_input("Pekerjaan Penjamin")
        al_pjm = st.text_area("Alamat Penjamin")
        h_pjm = c1.text_input("Hubungan Penjamin")
        hp_pjm = c2.text_input("No HP Penjamin")
        l_pjm = st.text_area("Kelayakan Penjamin")

    submit = st.form_submit_button("🚀 GENERATE LITMAS")

# --- 3. PROSES PEMBUATAN DOKUMEN ---
if submit:
    if not os.path.exists("template.docx"):
        st.error("File 'template.docx' tidak ditemukan!")
    else:
        try:
            with st.spinner("Sedang memproses dokumen dan analisis AI..."):
                # AI Analysis
                analisis_ai = generate_analisis_ai({
                    'nama_klien': nama_klien, 'perkara': perkara, 'pasal': pasal,
                    'latarbelakang_tp': latarbelakang_tp, 'kronologi': kronologi,
                    'keadaan_korban': keadaan_korban, 'sikap_klien': sikap_klien,
                    'hasil_rri': hasil_rri, 'hasil_kriminogenik': hasil_kriminogenik,
                    'bina_pribadi': bina_pribadi, 'bina_mandiri': bina_mandiri, 'program': program
                })

                # Persiapan Data Saudara
                list_saudara = []
                for _, row in saudara_editor.iterrows():
                    list_saudara.append({'no': row['No'], 'nama': row['Nama'], 'jk': row['JK'], 'usia': row['Usia'], 'pend': row['Pendidikan'], 'kerja': row['Pekerjaan'], 'ket': row['Keterangan']})

                # MAPPING SEMUA TAG
                context = {
                    'asal_upt': asal_upt, 'no_reg_litmas': no_reg_litmas, 'no_surat': no_surat,
                    'tgl_surat': format_tgl(tgl_surat), 'perihal': perihal, 'program': program,
                    'tgl_tpp': format_tgl(tgl_tpp), 'tgl_1per3': format_tgl(tgl_1per3),
                    'tgl_1per2': format_tgl(tgl_1per2), 'tgl_2per3': format_tgl(tgl_2per3),
                    'ekspirasi': ekspirasi, 'vonis': vonis, 'analisis_pk': analisis_ai, 'tgl_buat': format_tgl(datetime.now()),
                    
                    'nama_klien': nama_klien, 'no_reg': no_reg, 'ttl_klien': ttl_klien, 'jk_klien': jk_klien,
                    'suku_klien': suku_klien, 'warga_negara': warga_negara, 'agama_klien': agama_klien,
                    'pend_klien': pend_klien, 'sd': sd, 'smp': smp, 'sma': sma, 'perguruan_tinggi': pt,
                    'minat_bakat': minat_bakat, 'kerja_klien': kerja_klien, 'alamat_klien': alamat_klien,
                    'ciri_klien': ciri_klien, 'status_nikah_klien': status_nikah, 'riwayat_nikah': riwayat_nikah,
                    'riwayat_rokok': r_rokok, 'riwayat_alkohol': r_alkohol, 'riwayat_napza': r_napza,
                    
                    'perkara': perkara, 'pasal': pasal, 'latarbelakang_tp': latarbelakang_tp,
                    'kronologi': kronologi, 'keadaan_korban': keadaan_korban, 'sikap_klien': sikap_klien,
                    'bina_pribadi': bina_pribadi, 'bina_mandiri': bina_mandiri,
                    'hasil_rri': hasil_rri, 'hasil_kriminogenik': hasil_kriminogenik,
                    
                    'nama_ayah': n_ayah, 'ttl_ayah': ttl_ayah, 'agama_ayah': ag_ayah,
                    'suku_ayah': s_ayah, 'pend_ayah': p_ayah, 'kerja_ayah': k_ayah, 'alamat_ayah': al_ayah, 'ket_ayah': ket_ayah,
                    
                    'nama_ibu': n_ibu, 'ttl_ibu': ttl_ibu, 'agama_ibu': ag_ibu,
                    'suku_ibu': s_ibu, 'pend_ibu': p_ibu, 'kerja_ibu': k_ibu, 'alamat_ibu': al_ibu, 'ket_ibu': ket_ibu,
                    
                    'nama_pjm': n_pjm, 'ttl_pjm': ttl_pjm, 'agama_pjm': ag_pjm,
                    'suku_pjm': s_pjm, 'pend_pjm': p_pjm, 'kerja_pjm': k_pjm, 'alamat_pjm': al_pjm,
                    'hub_pjm': h_pjm, 'no_hp_pjm': hp_pjm, 'layak_pjm': l_pjm,
                    
                    'saudara': list_saudara
                }

                doc = DocxTemplate("template.docx")
                doc.render(context)
                
                output = io.BytesIO()
                doc.save(output)
                output.seek(0)

                st.success("✅ Berhasil! Silakan download dokumen di bawah ini.")
                st.download_button(label="📥 DOWNLOAD LITMAS", data=output, file_name=f"Litmas_{nama_klien}.docx")
        except Exception as e:
            st.error(f"Terjadi kesalahan teknis: {str(e)}")
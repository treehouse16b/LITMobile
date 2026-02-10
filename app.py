import streamlit as st
from docxtpl import DocxTemplate
import io
import os
import pandas as pd
from datetime import datetime
import google.generativeai as genai

# --- 1. KONFIGURASI AI ---
API_KEY = "AIzaSy..." # MASUKKAN API KEY ANDA DI SINI

def generate_litmas_parts(data):
    try:
        genai.configure(api_key=API_KEY)
        model = genai.GenerativeModel('gemini-1.5-flash')
        
        prompt = f"""
        Anda adalah Pembimbing Kemasyarakatan (PK) Ahli. Buatlah laporan Litmas dalam 3 bagian.
        DATA: Klien {data['nama_klien']}, Kasus {data['perkara']} {data['pasal']}, RRI: {data['hasil_rri']}, 
        Pembinaan: {data['bina_pribadi']} & {data['bina_mandiri']}, Respon: {data['respon_masyarakat']}.
        
        TUGAS:
        1. ANALISIS PK (Bab XI): Gunakan dasar teori (RNR/Planned Behavior). Bahas Sikap, Hasil Pembinaan, Penerimaan Masyarakat, dan Kelayakan.
        2. KESIMPULAN: Layak/Tidak program {data['program']}.
        3. REKOMENDASI (Bab XIII): Saran konkret pengawasan.

        Pisahkan dengan penanda ###ANALISIS###, ###KESIMPULAN###, ###REKOMENDASI###.
        """
        response = model.generate_content(prompt)
        full_text = response.text
        parts = {'analisis': full_text, 'kesimpulan': "", 'rekomendasi': ""}
        if "###ANALISIS###" in full_text and "###KESIMPULAN###" in full_text and "###REKOMENDASI###" in full_text:
            parts['analisis'] = full_text.split("###ANALISIS###")[1].split("###KESIMPULAN###")[0].strip()
            parts['kesimpulan'] = full_text.split("###KESIMPULAN###")[1].split("###REKOMENDASI###")[0].strip()
            parts['rekomendasi'] = full_text.split("###REKOMENDASI###")[1].strip()
        return parts
    except Exception as e:
        return {'analisis': f"Gagal AI: {e}", 'kesimpulan': "Manual", 'rekomendasi': "Manual"}

def format_tgl(tgl_obj):
    if not tgl_obj: return ""
    bulan_indo = {1: "Januari", 2: "Februari", 3: "Maret", 4: "April", 5: "Mei", 6: "Juni", 
                  7: "Juli", 8: "Agustus", 9: "September", 10: "Oktober", 11: "November", 12: "Desember"}
    return f"{tgl_obj.day} {bulan_indo[tgl_obj.month]} {tgl_obj.year}"

# --- 2. ANTARMUKA (UI) ---
st.set_page_config(page_title="Litmas Pro AI 2026", layout="wide")
st.title("📂 Generator Litmas Terintegrasi (Versi Full)")

with st.form("litmas_form"):
    # ADMINISTRASI
    st.header("🏢 I. Administrasi")
    c1, c2, c3 = st.columns(3)
    asal_upt = c1.text_input("Asal UPT")
    no_reg_litmas = c2.text_input("No. Reg LITMAS")
    no_surat = c3.text_input("No. Surat Permintaan")
    tgl_surat = c1.date_input("Tgl Surat")
    perihal = c2.text_input("Perihal")
    program = c3.selectbox("Program", ["PB", "CB", "CMB", "Asimilasi"])
    
    st.subheader("🗓️ Tahapan Pidana")
    c1, c2, c3, c4, c5 = st.columns(5)
    tgl_tpp = c1.date_input("Tgl TPP")
    tgl_1per3 = c2.date_input("Tgl 1/3")
    tgl_1per2 = c3.date_input("Tgl 1/2")
    tgl_2per3 = c4.date_input("Tgl 2/3")
    ekspirasi = c5.text_input("Tgl Ekspirasi (Bebas)")

    # HUKUM
    st.header("⚖️ II. Data Hukum")
    c1, c2, c3 = st.columns(3)
    perkara = c1.text_input("Perkara")
    pasal = c2.text_input("Pasal")
    vonis = c3.text_input("Vonis")
    latarbelakang_tp = st.text_area("Latar Belakang TP")
    kronologi = st.text_area("Kronologi Kejadian")
    keadaan_korban = st.text_area("Keadaan Korban")

    # KLIEN
    st.header("👤 III. Data Klien")
    c1, c2, c3 = st.columns(3)
    nama_klien = c1.text_input("Nama Klien")
    no_reg = c2.text_input("No. Reg. Bin")
    ttl_klien = c3.text_input("TTL Klien")
    jk_klien = c1.selectbox("Jenis Kelamin", ["Laki-laki", "Perempuan"])
    suku_klien = c2.text_input("Suku")
    warga_negara = c3.text_input("Warga Negara", "Indonesia")
    agama_klien = c1.text_input("Agama")
    pend_Klien = c2.text_input("Pendidikan Terakhir")
    kerja_klien = c3.text_input("Pekerjaan")
    alamat_klien = st.text_area("Alamat Klien")
    ciri_klien = st.text_input("Ciri Khusus")
    status_nikah_klien = st.text_input("Status Nikah")
    riwayat_nikah = st.text_area("Riwayat Pernikahan")
    
    st.subheader("🎓 Riwayat Pendidikan")
    c1, c2, c3, c4 = st.columns(4)
    sd = c1.text_input("SD")
    smp = c2.text_input("SMP")
    sma = c3.text_input("SMA")
    perguruan_tinggi = c4.text_input("Perguruan Tinggi")

    st.subheader("🚬 Kebiasaan & Asesmen")
    c1, c2, c3, c4 = st.columns(4)
    riwayat_rokok = c1.text_input("Rokok")
    riwayat_alkohol = c2.text_input("Alkohol")
    riwayat_napza = c3.text_input("NAPZA")
    minat_bakat = c4.text_input("Minat Bakat")
    
    sikap_klien = st.text_area("Sikap Selama Pembinaan")
    bina_pribadi = st.text_area("Pembinaan Kepribadian")
    bina_mandiri = st.text_area("Pembinaan Kemandirian")
    
    c1, c2 = st.columns(2)
    hasil_rri = c1.text_input("Hasil RRI")
    hasil_kriminogenik = c2.text_input("Hasil Kriminogenik")

    # KELUARGA
    st.header("👨‍👩‍👦 IV. Keluarga & Penjamin")
    st.subheader("👨‍👩‍👧‍👦 Tabel Susunan Keluarga (Saudara)")
    df_saudara = pd.DataFrame([{"No": 1, "Nama": "", "JK": "L", "Usia": "", "Pendidikan": "", "Pekerjaan": "", "Keterangan": ""}])
    saudara_editor = st.data_editor(df_saudara, num_rows="dynamic")

    with st.expander("Data Orang Tua (Ayah & Ibu)"):
        c1, c2 = st.columns(2)
        with c1:
            st.write("**Ayah**")
            nama_ayah = st.text_input("Nama Ayah")
            ttl_ayah = st.text_input("TTL Ayah")
            agama_ayah = st.text_input("Agama Ayah")
            suku_ayah = st.text_input("Suku Ayah")
            pend_ayah = st.text_input("Pend Ayah")
            kerja_ayah = st.text_input("Kerja Ayah")
            alamat_ayah = st.text_area("Alamat Ayah")
            ket_ayah = st.text_input("Ket Ayah")
        with c2:
            st.write("**Ibu**")
            nama_ibu = st.text_input("Nama Ibu")
            ttl_ibu = st.text_input("TTL Ibu")
            agama_ibu = st.text_input("Agama Ibu")
            suku_ibu = st.text_input("Suku Ibu")
            pend_ibu = st.text_input("Pend Ibu")
            kerja_ibu = st.text_input("Kerja Ibu")
            alamat_ibu = st.text_area("Alamat Ibu")
            ket_ibu = st.text_input("Ket Ibu")

    with st.expander("Data Penjamin & Masyarakat"):
        c1, c2, c3 = st.columns(3)
        nama_pjm = c1.text_input("Nama Penjamin")
        ttl_pjm = c2.text_input("TTL Penjamin")
        agama_pjm = c3.text_input("Agama Penjamin")
        suku_pjm = c1.text_input("Suku Penjamin")
        pend_pjm = c2.text_input("Pend Penjamin")
        kerja_pjm = c3.text_input("Kerja Penjamin")
        alamat_pjm = st.text_area("Alamat Penjamin")
        hub_pjm = c1.text_input("Hubungan Penjamin")
        no_hp_pjm = c2.text_input("HP Penjamin")
        layak_pjm = st.text_area("Hasil Kelayakan (Masyarakat/Korban)")
        respon_masyarakat = st.text_area("Respon Masyarakat (Untuk AI)")
        tanggapan_korban = st.text_area("Tanggapan Korban (Untuk AI)")

    submit = st.form_submit_button("🚀 GENERATE LITMAS LENGKAP")

# --- 3. PROSES DATA ---
if submit:
    with st.spinner("AI sedang menyusun laporan..."):
        hasil_ai = generate_litmas_parts({
            'nama_klien': nama_klien, 'perkara': perkara, 'pasal': pasal, 'kronologi': kronologi,
            'sikap_klien': sikap_klien, 'bina_pribadi': bina_pribadi, 'bina_mandiri': bina_mandiri,
            'hasil_rri': hasil_rri, 'hasil_kriminogenik': hasil_kriminogenik,
            'respon_masyarakat': respon_masyarakat, 'tanggapan_korban': tanggapan_korban, 'program': program
        })
        
        list_saudara = []
        for _, row in saudara_editor.iterrows():
            list_saudara.append({'no': row['No'], 'nama': row['Nama'], 'jk': row['JK'], 'usia': row['Usia'], 'pend': row['Pendidikan'], 'kerja': row['Pekerjaan'], 'ket': row['Keterangan']})

        context = {
            'asal_upt': asal_upt, 'no_reg_litmas': no_reg_litmas, 'no_surat': no_surat, 'tgl_surat': format_tgl(tgl_surat),
            'perihal': perihal, 'program': program, 'tgl_tpp': format_tgl(tgl_tpp), 'tgl_1per3': format_tgl(tgl_1per3),
            'tgl_1per2': format_tgl(tgl_1per2), 'tgl_2per3': format_tgl(tgl_2per3), 'ekspirasi': ekspirasi, 'vonis': vonis,
            'nama_klien': nama_klien, 'no_reg': no_reg, 'ttl_klien': ttl_klien, 'jk_klien': jk_klien, 'suku_klien': suku_klien,
            'warga_negara': warga_negara, 'agama_klien': agama_klien, 'pend_Klien': pend_Klien, 'sd': sd, 'smp': smp, 'sma': sma,
            'perguruan_tinggi': perguruan_tinggi, 'minat_bakat': minat_bakat, 'kerja_klien': kerja_klien, 'alamat_klien': alamat_klien,
            'ciri_klien': ciri_klien, 'status_nikah_klien': status_nikah_klien, 'riwayat_nikah': riwayat_nikah,
            'riwayat_rokok': riwayat_rokok, 'riwayat_alkohol': riwayat_alkohol, 'riwayat_napza': riwayat_napza,
            'perkara': perkara, 'pasal': pasal, 'latarbelakang_tp': latarbelakang_tp, 'kronologi': kronologi,
            'keadaan_korban': keadaan_korban, 'sikap_klien': sikap_klien, 'bina_pribadi': bina_pribadi, 'bina_mandiri': bina_mandiri,
            'hasil_rri': hasil_rri, 'hasil_kriminogenik': hasil_kriminogenik,
            'nama_ayah': nama_ayah, 'ttl_ayah': ttl_ayah, 'agama_ayah': agama_ayah, 'suku_ayah': suku_ayah,
            'pend_ayah': pend_ayah, 'kerja_ayah': kerja_ayah, 'alamat_ayah': alamat_ayah, 'ket_ayah': ket_ayah,
            'nama_ibu': nama_ibu, 'ttl_ibu': ttl_ibu, 'agama_ibu': agama_ibu, 'suku_ibu': suku_ibu,
            'pend_ibu': pend_ibu, 'kerja_ibu': kerja_ibu, 'alamat_ibu': alamat_ibu, 'ket_ibu': ket_ibu,
            'nama_pjm': nama_pjm, 'ttl_pjm': ttl_pjm, 'agama_pjm': agama_pjm, 'suku_pjm': suku_pjm,
            'pend_pjm': pend_pjm, 'kerja_pjm': kerja_pjm, 'alamat_pjm': alamat_pjm, 'hub_pjm': hub_pjm,
            'no_hp_pjm': no_hp_pjm, 'layak_pjm': layak_pjm,
            'saudara': list_saudara, 'analisis_pk': hasil_ai['analisis'], 'kesimpulan': hasil_ai['kesimpulan'], 
            'rekomendasi_pk': hasil_ai['rekomendasi'], 'tgl_buat': format_tgl(datetime.now())
        }

        try:
            doc = DocxTemplate("template.docx")
            doc.render(context)
            bio = io.BytesIO()
            doc.save(bio)
            bio.seek(0)
            st.success("✅ Dokumen Siap!")
            st.download_button("📥 DOWNLOAD WORD", bio, f"Litmas_Full_{nama_klien}.docx")
        except Exception as e:
            st.error(f"Gagal: {e}")

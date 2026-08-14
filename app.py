import streamlit as st
import pandas as pd
import datetime
import io
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from google import genai

# ==========================================
# 1. SAYFA VE GÜVENLİK AYARLARI
# ==========================================
st.set_page_config(page_title="R&D LIMS Intelligence Cloud", layout="wide", page_icon="🧪")

# Kullanıcı Hesapları (VBA'deki Sistem_Kullanicilari mantığı)
USERS = {
    "abban": "1234",      # Kullanıcı 1 (Kullanıcı adı: abban, Şifre: 1234)
    "partner": "5678"     # Kullanıcı 2
}

if "logged_in" not in st.session_state:
    st.session_state["logged_in"] = False
if "username" not in st.session_state:
    st.session_state["username"] = ""

# --- GİRİŞ PANELDEN (LOGIN) ---
if not st.session_state["logged_in"]:
    st.markdown("<h2 style='text-align: center; color: #10365C;'>🧪 R&D Intelligence — Giriş Paneli</h2>", unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        user_input = st.text_input("Kullanıcı Adı")
        pass_input = st.text_input("Şifre", type="password")
        if st.button("Sisteme Giriş Yap"):
            if user_input.lower() in USERS and USERS[user_input.lower()] == pass_input:
                st.session_state["logged_in"] = True
                st.session_state["username"] = user_input.capitalize()
                st.success("Giriş Başarılı!")
                st.rerun()
            else:
                st.error("Hatalı Kullanıcı Adı veya Şifre!")
    st.stop()

# ==========================================
# 2. ANA UYGULAMA (GİRİŞ YAPILDIKTAN SONRA)
# ==========================================
st.sidebar.title(f"👤 Kullanıcı: {st.session_state['username']}")
if st.sidebar.button("Güvenli Çıkış"):
    st.session_state["logged_in"] = False
    st.rerun()

st.sidebar.markdown("---")
secim = st.sidebar.radio("İşlem Modülü:", [
    "🧪 Formülasyon & Analiz Tezgahı", 
    "📄 CoA / CoC / SOP Belge Üretimi", 
    "🤖 Gemini AI Ar-Ge Asistanı"
])

# ------------------------------------------
# MODÜL 1: FORMÜLASYON & ANALİZ TEZGAHI
# ------------------------------------------
if secim == "🧪 Formülasyon & Analiz Tezgahı":
    st.markdown("<h3 style='color:#10365C;'>Formülasyon Geliştirme ve Analiz Motoru</h3>", unsafe_allow_html=True)
    
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        tip = st.selectbox("Formülasyon Tipi", ["LIQ", "NPK", "GRA", "MIC", "SC"])
    with c2:
        talep_no = st.text_input("Talep No", "MSTR01-260814-1600")
    with c3:
        urun_adi = st.text_input("Ürün Adı", "AMINOSOL+TE")
    with c4:
        yogunluk = st.number_input("d (Yoğunluk g/cm³)", min_value=0.1, max_value=2.5, value=1.25, step=0.01)

    st.markdown("---")
    st.write("##### Hammadde ve Reçete Oranları (% w/w)")
    
    # Varsayılan Örnek Reçete
    default_data = pd.DataFrame([
        {"Hammadde Adı": "Su", "Miktar (%)": 45.00},
        {"Hammadde Adı": "Amino Asit Kompleks", "Miktar (%)": 30.00},
        {"Hammadde Adı": "Çinko Sülfat Monohidrat", "Miktar (%)": 25.00}
    ])
    
    edited_df = st.data_editor(default_data, num_rows="dynamic", use_container_width=True)
    
    toplam = edited_df["Miktar (%)"].sum()
    
    # 100 kg Takip Barı
    if abs(toplam - 100.0) < 0.001:
        st.success(f"✅ TOPLAM: {toplam:.2f} kg — Formülasyon Dengede!")
    elif toplam < 100.0:
        st.warning(f"⚠️ TOPLAM: {toplam:.2f} kg — Kalan Eksik: {100.0 - toplam:.2f} kg")
    else:
        st.error(f"❌ TOPLAM: {toplam:.2f} kg — Fazla Miktar: {toplam - 100.0:.2f} kg")

    if st.button("Formül Kodu Oluştur ve Arşivle"):
        zaman = datetime.datetime.now().strftime("%y%m%d-%H%M")
        kod = f"{tip}-{zaman}R0 | Rev-00"
        st.info(f"Oluşturulan Otomatik Formül Kodu: **{kod}**")

# ------------------------------------------
# MODÜL 2: GEMINI AI ASİSTANI
# ------------------------------------------
elif secim == "🤖 Gemini AI Ar-Ge Asistanı":
    st.markdown("<h3 style='color:#10365C;'>Gemini AI Ar-Ge & Kimya Asistanı</h3>", unsafe_allow_html=True)
    
    api_key = st.text_input("Gemini API Key:", type="password")
    prompt = st.text_area("Ar-Ge Sorunuz veya Reçete Talebiniz:", placeholder="Örn: Sıvı gübrede çinko sülfat ile potasyum humat çökelme yapar mı?")
    
    if st.button("Soru Sor"):
        if not api_key:
            st.error("Lütfen API Key giriniz.")
        else:
            try:
                client = genai.Client(api_key=api_key)
                res = client.models.generate_content(
                    model='gemini-2.5-flash',
                    contents=prompt
                )
                st.markdown("### 🤖 Asistan Cevabı:")
                st.write(res.text)
            except Exception as e:
                st.error(f"Hata: {e}")
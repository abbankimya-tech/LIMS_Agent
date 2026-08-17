import streamlit as st
import pandas as pd
import datetime
import io
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

# ==========================================
# 1. SAYFA VE GÜVENLİK AYARLARI
# ==========================================
st.set_page_config(page_title="R&D LIMS Intelligence", layout="wide", page_icon="🧪")

USERS = {
    "abban": "1234",
    "partner": "5678"
}

if "logged_in" not in st.session_state:
    st.session_state["logged_in"] = False
if "username" not in st.session_state:
    st.session_state["username"] = ""

if not st.session_state["logged_in"]:
    st.markdown("<h2 style='text-align: center; color: #10365C;'>🧪 R&D Intelligence — Giriş Paneli</h2>", unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        user_input = st.text_input("Kullanıcı Adı")
        pass_input = st.text_input("Şifre", type="password")
        if st.button("Sisteme Giriş Yap", use_container_width=True):
            if user_input.lower() in USERS and USERS[user_input.lower()] == pass_input:
                st.session_state["logged_in"] = True
                st.session_state["username"] = user_input.capitalize()
                st.rerun()
            else:
                st.error("Hatalı Kullanıcı Adı veya Şifre!")
    st.stop()

# ==========================================
# 2. HAMMADDELER & ANALİZ VERİTABANI
# ==========================================
# PDF'ten Okunan Hammadde Veri Tablosu
RAW_MATERIALS_DATA = [
    {"Kod": "10100004", "Ad": "Amino Asit", "Total N": 6.0, "Organic Matter": 95.0, "Free Aminoacids": 24.0},
    {"Kod": "10100008", "Ad": "Amonyum Nitrat", "Total N": 33.0, "NH4-N": 16.5, "NO3-N": 16.5},
    {"Kod": "10100233", "Ad": "Amonyum Polifosfat 18-58-0", "Total N": 18.0, "NH4-N": 18.0, "Total P2O5": 58.0},
    {"Kod": "10100009", "Ad": "Amonyum Sülfat", "Total N": 21.0, "NH4-N": 21.0, "SO3": 60.0},
    {"Kod": "10100021", "Ad": "Borik Asit", "Water Soluble B": 17.0},
    {"Kod": "10100028", "Ad": "Çinko Edta %15", "Water Soluble Zn": 15.0},
    {"Kod": "10100033", "Ad": "Çinko Sülfat Monohidrat", "SO3": 42.9, "Water Soluble Zn": 35.0},
    {"Kod": "10100036", "Ad": "Demir Edta %13", "Water Soluble Fe": 13.0},
    {"Kod": "10100037", "Ad": "Demir Sülfat Heptahidrat", "SO3": 27.3, "Water Soluble Fe": 19.0},
    {"Kod": "10100038", "Ad": "Demir Sülfat Monohidrat", "SO3": 40.1, "Water Soluble Fe": 28.0},
    {"Kod": "10100055", "Ad": "Fosforik Asit", "Total P2O5": 61.0},
    {"Kod": "10100056", "Ad": "Fosforoz Asit", "Total P2O5": 85.0},
    {"Kod": "10100066", "Ad": "Kalsiyum Nitrat", "Total N": 15.5, "NO3-N": 14.5, "CaO": 26.5},
    {"Kod": "10100078", "Ad": "Magnezyum Nitrat", "Total N": 10.0, "NO3-N": 10.0, "MgO": 15.0},
    {"Kod": "10100083", "Ad": "Magnezyum Sülfat Heptahidrat", "MgO": 15.0, "SO3": 30.0},
    {"Kod": "10100089", "Ad": "Mangan Sülfat Monohidrat", "SO3": 46.6, "Water Soluble Mn": 32.0},
    {"Kod": "10100090", "Ad": "MAP (Mono Amonyum Fosfat)", "Total N": 12.0, "NH4-N": 12.0, "Total P2O5": 61.0},
    {"Kod": "10100098", "Ad": "MKP (Mono Potasyum Fosfat)", "Total P2O5": 52.0, "K2O": 34.0},
    {"Kod": "10100118", "Ad": "Potasyum Hidroksit (Flake)", "K2O": 75.0},
    {"Kod": "10100123", "Ad": "Potasyum Nitrat (LP)", "Total N": 13.0, "NO3-N": 13.0, "K2O": 46.0},
    {"Kod": "10100126", "Ad": "Potasyum Sülfat", "K2O": 51.0, "SO3": 43.0, "S": 17.4},
    {"Kod": "10100128", "Ad": "Protein Hidrolizati 70", "Total N": 8.0, "Organic Matter": 55.0, "Free Aminoacids": 22.0},
    {"Kod": "10100133", "Ad": "Sitrik Asit Anhidrat", "Organic Matter": 98.0},
    {"Kod": "10100143", "Ad": "Su", "Total N": 0.0},
    {"Kod": "10100161", "Ad": "Üre", "Total N": 46.0, "NH2-N": 46.0},
    {"Kod": "kodu yok", "Ad": "Üre Fosfat", "Total N": 17.0, "NH2-N": 17.0, "Total P2O5": 44.0}
]

df_raw_db = pd.DataFrame(RAW_MATERIALS_DATA).fillna(0.0)
HAMMADDELER_LISTESI = df_raw_db["Ad"].tolist()

if "FORMULLER_DB" not in st.session_state:
    st.session_state["FORMULLER_DB"] = []

# ==========================================
# 3. YAN MENÜ
# ==========================================
st.sidebar.title(f"👤 Kullanıcı: {st.session_state['username']}")
if st.sidebar.button("Güvenli Çıkış"):
    st.session_state["logged_in"] = False
    st.rerun()

st.sidebar.markdown("---")
secim = st.sidebar.radio("İşlem Modülü:", [
    "🧪 Formülasyon Tezgahı", 
    "📄 Belge Üretimi (CoA/CoC/SOP)", 
    "📚 Reçete Veritabanı (FORMULLER_DB)",
    "⚙️ Hammadde & Analiz Veritabanı"
])

# ==========================================
# MODÜL 1: FORMÜLASYON TEZGAHI
# ==========================================
if secim == "🧪 Formülasyon Tezgahı":
    st.markdown("<h3 style='color:#10365C;'>🧪 Formülasyon Oluşturma & Analiz Tezgahı</h3>", unsafe_allow_html=True)
    
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        tip = st.selectbox("Formülasyon Tipi", ["LIQ", "NPK", "GRA", "MIC", "SC"])
    with c2:
        talep_no = st.text_input("Talep / Müşteri No", "MSTR01-260816-2100")
    with c3:
        urun_adi = st.text_input("Ürün Ticari Adı", "AMINOSOL+TE")
    with c4:
        yogunluk = st.number_input("Yoğunluk d (g/cm³)", min_value=0.100, max_value=2.500, value=1.250, step=0.001, format="%.3f")

    c5, c6 = st.columns(2)
    with c5:
        ph_degeri = st.number_input("pH (1/10 Çözeltide)", min_value=0.0, max_value=14.0, value=6.5, step=0.1)
    with c6:
        firma = st.selectbox("Üretici / Marka Şablonu", ["Altıntar Tarım A.Ş.", "Unifarm Chemical"])

    st.markdown("---")
    st.write("##### 📋 Hammadde Reçete Girişi (% w/w)")
    
    if "df_recipe" not in st.session_state:
        st.session_state["df_recipe"] = pd.DataFrame([
            {"Hammadde Adı": "Su", "Miktar (% w/w)": 45.00},
            {"Hammadde Adı": "Protein Hidrolizati 70", "Miktar (% w/w)": 30.00},
            {"Hammadde Adı": "Çinko Sülfat Monohidrat", "Miktar (% w/w)": 25.00}
        ])

    edited_df = st.data_editor(
        st.session_state["df_recipe"],
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "Hammadde Adı": st.column_config.SelectboxColumn(
                "Hammadde Adı",
                options=HAMMADDELER_LISTESI,
                required=True
            ),
            "Miktar (% w/w)": st.column_config.NumberColumn(
                "Miktar (% w/w)",
                min_value=0.0,
                max_value=100.0,
                step=0.1,
                format="%.2f %%"
            )
        }
    )
    
    toplam_w_w = edited_df["Miktar (% w/w)"].sum() if not edited_df.empty else 0.0
    
    if abs(toplam_w_w - 100.0) < 0.001:
        st.success(f"✅ TOPLAM BİLEŞİM: %{toplam_w_w:.2f} (w/w) — Reçete Tamamlandı!")
    elif toplam_w_w < 100.0:
        st.warning(f"⚠️ TOPLAM BİLEŞİM: %{toplam_w_w:.2f} (w/w) — Kalan Eksik: %{100.0 - toplam_w_w:.2f}")
    else:
        st.error(f"❌ TOPLAM BİLEŞİM: %{toplam_w_w:.2f} (w/w) — Fazla Miktar: %{toplam_w_w - 100.0:.2f}")

    # --- GARANTİ EDİLEN İÇERİK OTOMATİK HESAPLAMA MOTORU ---
    st.markdown("---")
    st.write("##### 📊 Garanti Edilen İçerik / Hesaplanan Analiz Değerleri")
    
    if not edited_df.empty:
        # Analiz Sütunlarını Hesapla
        totals_ww = {}
        for idx, row in edited_df.iterrows():
            h_ad = row["Hammadde Adı"]
            h_miktar = row["Miktar (% w/w)"]
            
            # DB'den hammaddeyi bul
            match = df_raw_db[df_raw_db["Ad"] == h_ad]
            if not match.empty:
                h_row = match.iloc[0]
                for col in df_raw_db.columns:
                    if col not in ["Kod", "Ad"]:
                        val = float(h_row[col]) if col in h_row else 0.0
                        if val > 0:
                            contrib = (h_miktar * val) / 100.0
                            totals_ww[col] = totals_ww.get(col, 0.0) + contrib

        if totals_ww:
            analysis_summary = []
            for param, val_ww in totals_ww.items():
                if val_ww > 0.001:
                    val_wv = val_ww * yogunluk
                    val_gl = val_wv * 10.0
                    analysis_summary.append({
                        "Analiz Parametresi": param,
                        "w/w % (Ağırlıkça)": f"{val_ww:.2f} %",
                        "w/v % (Hacimce)": f"{val_wv:.2f} %",
                        "g/L (Derişim)": f"{val_gl:.2f} g/L"
                    })
            
            df_analysis_out = pd.DataFrame(analysis_summary)
            st.dataframe(df_analysis_out, use_container_width=True)
        else:
            st.info("Seçilen hammaddelerde tanımlı analiz değeri bulunmuyor.")

# ==========================================
# MODÜL 4: HAMMADDE YÖNETİMİ
# ==========================================
elif secim == "⚙️ Hammadde & Analiz Veritabanı":
    st.markdown("<h3 style='color:#10365C;'>⚙️ Hammadde & Analiz Değerleri Yönetim Paneli</h3>", unsafe_allow_html=True)
    st.write("Bu panelden sistemdeki tüm hammaddeleri ve analiz oranlarını inceleyebilir, yeni hammaddeler ekleyebilirsiniz.")
    st.dataframe(df_raw_db, use_container_width=True)

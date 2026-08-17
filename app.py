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
# 2. SABİT VERİLER VE VERİTABANI HAFIZASI
# ==========================================

# Müşteri Veritabanı (MUSTERI_DATA)
MUSTERI_DATA = [
    {"Musteri": "AgroCorp GmbH", "Kod": "AGRO01", "Ulke": "Almanya"},
    {"Musteri": "Altıntar Tarım A.Ş.", "Kod": "ALT01", "Ulke": "Türkiye"},
    {"Musteri": "Unifarm Chemical", "Kod": "UNI01", "Ulke": "Türkiye"},
    {"Musteri": "GreenFarm Agro", "Kod": "GRN02", "Ulke": "İspanya"},
    {"Musteri": "BioPlant SRL", "Kod": "BIO03", "Ulke": "İtalya"}
]

# TALEPLER Veritabanı İlklendirme
if "TALEPLER_DB" not in st.session_state:
    st.session_state["TALEPLER_DB"] = pd.DataFrame([
        {
            "Talep No": "AGRO01-260817-1030",
            "Kayıt Tarihi": "17.08.2026 10:30",
            "Ülke": "Almanya",
            "Müşteri / Firma Adı": "AgroCorp GmbH",
            "İlgili Kişi": "Hans Müller",
            "Talep Tipi": "Formülasyon",
            "Ürün İsmi": "AMINOSOL+TE",
            "Durum": "Bekliyor",
            "Üretilen Kod": "-",
            "Açıklama / Notlar": "Sıvı çinko aminoasit şelatı talebi"
        },
        {
            "Talep No": "ALT01-260816-1415",
            "Kayıt Tarihi": "16.08.2026 14:15",
            "Ülke": "Türkiye",
            "Müşteri / Firma Adı": "Altıntar Tarım A.Ş.",
            "İlgili Kişi": "Ahmet Bey",
            "Talep Tipi": "Formülasyon Revizyonu",
            "Ürün İsmi": "HUMICAL LIQ",
            "Durum": "2 - Arşive Eklendi (Onay Bekliyor)",
            "Üretilen Kod": "LIQ-260816-1415R0",
            "Açıklama / Notlar": "pH 6.5 hedefli revizyon"
        }
    ])

# Formüller Veritabanı
if "FORMULLER_DB" not in st.session_state:
    st.session_state["FORMULLER_DB"] = []

# Active Talep Transfer Hafızası
if "active_talep_no" not in st.session_state:
    st.session_state["active_talep_no"] = "MSTR01-260817-2100"
if "active_urun_adi" not in st.session_state:
    st.session_state["active_urun_adi"] = "AMINOSOL+TE"

# Hammadde Veritabanı
if "df_raw_db" not in st.session_state:
    RAW_MATERIALS_DATA = [
        {"Kod": "10100004", "Ad": "Amino Asit", "Total N": 6.0, "Organic Matter": 95.0, "Free Aminoacids": 24.0, "Total P2O5": 0.0, "K2O": 0.0, "Water Soluble Zn": 0.0, "Water Soluble Fe": 0.0, "Water Soluble B": 0.0, "SO3": 0.0},
        {"Kod": "10100008", "Ad": "Amonyum Nitrat", "Total N": 33.0, "Organic Matter": 0.0, "Free Aminoacids": 0.0, "Total P2O5": 0.0, "K2O": 0.0, "Water Soluble Zn": 0.0, "Water Soluble Fe": 0.0, "Water Soluble B": 0.0, "SO3": 0.0},
        {"Kod": "10100233", "Ad": "Amonyum Polifosfat 18-58-0", "Total N": 18.0, "Organic Matter": 0.0, "Free Aminoacids": 0.0, "Total P2O5": 58.0, "K2O": 0.0, "Water Soluble Zn": 0.0, "Water Soluble Fe": 0.0, "Water Soluble B": 0.0, "SO3": 0.0},
        {"Kod": "10100009", "Ad": "Amonyum Sülfat", "Total N": 21.0, "Organic Matter": 0.0, "Free Aminoacids": 0.0, "Total P2O5": 0.0, "K2O": 0.0, "Water Soluble Zn": 0.0, "Water Soluble Fe": 0.0, "Water Soluble B": 0.0, "SO3": 60.0},
        {"Kod": "10100021", "Ad": "Borik Asit", "Total N": 0.0, "Organic Matter": 0.0, "Free Aminoacids": 0.0, "Total P2O5": 0.0, "K2O": 0.0, "Water Soluble Zn": 0.0, "Water Soluble Fe": 0.0, "Water Soluble B": 17.0, "SO3": 0.0},
        {"Kod": "10100028", "Ad": "Çinko Edta %15", "Total N": 0.0, "Organic Matter": 0.0, "Free Aminoacids": 0.0, "Total P2O5": 0.0, "K2O": 0.0, "Water Soluble Zn": 15.0, "Water Soluble Fe": 0.0, "Water Soluble B": 0.0, "SO3": 0.0},
        {"Kod": "10100033", "Ad": "Çinko Sülfat Monohidrat", "Total N": 0.0, "Organic Matter": 0.0, "Free Aminoacids": 0.0, "Total P2O5": 0.0, "K2O": 0.0, "Water Soluble Zn": 35.0, "Water Soluble Fe": 0.0, "Water Soluble B": 0.0, "SO3": 42.9},
        {"Kod": "10100036", "Ad": "Demir Edta %13", "Total N": 0.0, "Organic Matter": 0.0, "Free Aminoacids": 0.0, "Total P2O5": 0.0, "K2O": 0.0, "Water Soluble Zn": 0.0, "Water Soluble Fe": 13.0, "Water Soluble B": 0.0, "SO3": 0.0},
        {"Kod": "10100055", "Ad": "Fosforik Asit", "Total N": 0.0, "Organic Matter": 0.0, "Free Aminoacids": 0.0, "Total P2O5": 61.0, "K2O": 0.0, "Water Soluble Zn": 0.0, "Water Soluble Fe": 0.0, "Water Soluble B": 0.0, "SO3": 0.0},
        {"Kod": "10100066", "Ad": "Kalsiyum Nitrat", "Total N": 15.5, "Organic Matter": 0.0, "Free Aminoacids": 0.0, "Total P2O5": 0.0, "K2O": 0.0, "Water Soluble Zn": 0.0, "Water Soluble Fe": 0.0, "Water Soluble B": 0.0, "SO3": 0.0},
        {"Kod": "10100090", "Ad": "MAP (Mono Amonyum Fosfat)", "Total N": 12.0, "Organic Matter": 0.0, "Free Aminoacids": 0.0, "Total P2O5": 61.0, "K2O": 0.0, "Water Soluble Zn": 0.0, "Water Soluble Fe": 0.0, "Water Soluble B": 0.0, "SO3": 0.0},
        {"Kod": "10100098", "Ad": "MKP (Mono Potasyum Fosfat)", "Total N": 0.0, "Organic Matter": 0.0, "Free Aminoacids": 0.0, "Total P2O5": 52.0, "K2O": 34.0, "Water Soluble Zn": 0.0, "Water Soluble Fe": 0.0, "Water Soluble B": 0.0, "SO3": 0.0},
        {"Kod": "10100118", "Ad": "Potasyum Hidroksit (Flake)", "Total N": 0.0, "Organic Matter": 0.0, "Free Aminoacids": 0.0, "Total P2O5": 0.0, "K2O": 75.0, "Water Soluble Zn": 0.0, "Water Soluble Fe": 0.0, "Water Soluble B": 0.0, "SO3": 0.0},
        {"Kod": "10100123", "Ad": "Potasyum Nitrat (LP)", "Total N": 13.0, "Organic Matter": 0.0, "Free Aminoacids": 0.0, "Total P2O5": 0.0, "K2O": 46.0, "Water Soluble Zn": 0.0, "Water Soluble Fe": 0.0, "Water Soluble B": 0.0, "SO3": 0.0},
        {"Kod": "10100126", "Ad": "Potasyum Sülfat", "Total N": 0.0, "Organic Matter": 0.0, "Free Aminoacids": 0.0, "Total P2O5": 0.0, "K2O": 51.0, "Water Soluble Zn": 0.0, "Water Soluble Fe": 0.0, "Water Soluble B": 0.0, "SO3": 43.0},
        {"Kod": "10100128", "Ad": "Protein Hidrolizati 70", "Total N": 8.0, "Organic Matter": 55.0, "Free Aminoacids": 22.0, "Total P2O5": 0.0, "K2O": 0.0, "Water Soluble Zn": 0.0, "Water Soluble Fe": 0.0, "Water Soluble B": 0.0, "SO3": 0.0},
        {"Kod": "10100143", "Ad": "Su", "Total N": 0.0, "Organic Matter": 0.0, "Free Aminoacids": 0.0, "Total P2O5": 0.0, "K2O": 0.0, "Water Soluble Zn": 0.0, "Water Soluble Fe": 0.0, "Water Soluble B": 0.0, "SO3": 0.0},
        {"Kod": "10100161", "Ad": "Üre", "Total N": 46.0, "Organic Matter": 0.0, "Free Aminoacids": 0.0, "Total P2O5": 0.0, "K2O": 0.0, "Water Soluble Zn": 0.0, "Water Soluble Fe": 0.0, "Water Soluble B": 0.0, "SO3": 0.0}
    ]
    st.session_state["df_raw_db"] = pd.DataFrame(RAW_MATERIALS_DATA).fillna(0.0)

df_raw_db = st.session_state["df_raw_db"]
HAMMADDELER_LISTESI = df_raw_db["Ad"].dropna().tolist()

# ==========================================
# 3. YAN MENÜ VE GEZİNTİ
# ==========================================
st.sidebar.title(f"👤 Kullanıcı: {st.session_state['username']}")
if st.sidebar.button("Güvenli Çıkış"):
    st.session_state["logged_in"] = False
    st.rerun()

st.sidebar.markdown("---")
# TALEPLER EN BAŞTA OLABİLECEK ŞEKİLDE SIRALANDI
secim = st.sidebar.radio("İşlem Modülü:", [
    "📋 Talepler Yönetimi",
    "🧪 Formülasyon Tezgahı", 
    "📄 Belge Üretimi (CoA/CoC/SOP)", 
    "📚 Reçete Veritabanı (FORMULLER_DB)",
    "⚙️ Hammadde & Analiz Veritabanı"
])

# ==========================================
# MODÜL 0: TALEPLER YÖNETİMİ (EN BAŞTA)
# ==========================================
if secim == "📋 Talepler Yönetimi":
    st.markdown("<h3 style='color:#10365C;'>📋 Müşteri Talepleri ve Ar-Ge Süreç Takibi</h3>", unsafe_allow_html=True)
    
    # 1. Yeni Talep Girişi (Expander Form)
    with st.expander("➕ Yeni Talep Oluştur (Müşteri Talebi Ekle)", expanded=False):
        c1, c2, c3 = st.columns(3)
        with c1:
            musteri_isimleri = [m["Musteri"] for m in MUSTERI_DATA]
            secilen_musteri = st.selectbox("Müşteri / Firma Seçiniz", musteri_isimleri)
            
            # Otomatik Bilgi Çekme (MÜŞTERİ_DATA'dan)
            m_info = next((m for m in MUSTERI_DATA if m["Musteri"] == secilen_musteri), {"Kod": "MSTR01", "Ulke": "Türkiye"})
            m_kod = m_info["Kod"]
            m_ulke = m_info["Ulke"]
            
        with c2:
            st.text_input("Ülke", value=m_ulke, disabled=True)
            # Otomatik Kod Üretici: MUSTERIKODU-YYMMDD-HHMM
            zaman_str = datetime.datetime.now().strftime("%y%m%d-%H%M")
            otomatik_talep_no = f"{m_kod}-{zaman_str}"
            st.text_input("Üretilen Otomatik Talep No", value=otomatik_talep_no, disabled=True)
            
        with c3:
            ilgili_kisi = st.text_input("İlgili Kişi", "Müşteri Yetkilisi")
            talep_tipi = st.selectbox("Talep Tipi", ["Formülasyon", "Formülasyon Revizyonu", "Hammadde QC", "Rakip Analizi", "CoA", "Üretim Prosesi", "Diğer"])

        c4, c5 = st.columns(2)
        with c4:
            urun_ismi = st.text_input("Talep Edilen Ürün İsmi", "AMINOSOL+TE")
        with c5:
            aciklama = st.text_area("Açıklama / Notlar", "Müşteri isteği ve teknik detaylar...")

        if st.button("💾 Talebi Sistem Kaydına Al", use_container_width=True):
            yeni_satir = {
                "Talep No": otomatik_talep_no,
                "Kayıt Tarihi": datetime.datetime.now().strftime("%d.%m.%Y %H:%M"),
                "Ülke": m_ulke,
                "Müşteri / Firma Adı": secilen_musteri,
                "İlgili Kişi": ilgili_kisi,
                "Talep Tipi": talep_tipi,
                "Ürün İsmi": urun_ismi,
                "Durum": "Bekliyor",
                "Üretilen Kod": "-",
                "Açıklama / Notlar": aciklama
            }
            st.session_state["TALEPLER_DB"] = pd.concat([pd.DataFrame([yeni_satir]), st.session_state["TALEPLER_DB"]], ignore_index=True)
            st.success(f"✅ Talep Başarıyla Kaydedildi! | **Talep No:** `{otomatik_talep_no}`")
            st.rerun()

    st.markdown("---")
    st.write("##### 📊 Mevcut Müşteri Talepleri Listesi")
    
    # Filtreleme ve Tablo
    df_talepler = st.session_state["TALEPLER_DB"]
    
    # Canlı Tablo Gösterimi
    st.dataframe(df_talepler, use_container_width=True)
    
    st.markdown("---")
    st.write("##### 🧪 Talep Üzerinde Çalışma / Formülasyona Geçiş")
    
    if not df_talepler.empty:
        s1, s2 = st.columns([3, 1])
        with s1:
            secilen_t_no = st.selectbox("Formülasyon Çalışması Başlatılacak Talebi Seçin:", df_talepler["Talep No"])
        with s2:
            st.write("")
            st.write("")
            if st.button("🚀 Bu Talebi Formüle Et", use_container_width=True):
                # Seçilen talebin ürün ismini bul
                row_match = df_talepler[df_talepler["Talep No"] == secilen_t_no].iloc[0]
                
                st.session_state["active_talep_no"] = secilen_t_no
                st.session_state["active_urun_adi"] = row_match["Ürün İsmi"]
                
                st.success(f"Talep `# {secilen_t_no}` Formülasyon Tezgahına aktarıldı! Lütfen sol menüden **🧪 Formülasyon Tezgahı** sekmesine geçiniz.")

# ==========================================
# MODÜL 1: FORMÜLASYON TEZGAHI
# ==========================================
elif secim == "🧪 Formülasyon Tezgahı":
    st.markdown("<h3 style='color:#10365C;'>🧪 Formülasyon Oluşturma & Analiz Tezgahı</h3>", unsafe_allow_html=True)
    
    # Aktif Talepten Veri Çekme
    default_talep = st.session_state.get("active_talep_no", "MSTR01-260817-2100")
    default_urun = st.session_state.get("active_urun_adi", "AMINOSOL+TE")
    
    st.info(f"📌 **Aktif Çalışılan Talep:** `{default_talep}` | **Ürün:** `{default_urun}`")
    
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        tip = st.selectbox("Formülasyon Tipi", ["LIQ", "NPK", "GRA", "MIC", "SC"])
    with c2:
        talep_no = st.text_input("Talep / Müşteri No", value=default_talep)
    with c3:
        urun_adi = st.text_input("Ürün Ticari Adı", value=default_urun)
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
        st.error(f"❌ TOPLAM BİLEŞİM: %{toplam_w_w:.2f} (w/w) — Fazla Miktar: %{toplam_w_w - toplam_w_w:.2f}")

    # Garanti Edilen İçerik Motoru
    st.markdown("---")
    st.write("##### 📊 Garanti Edilen İçerik / Hesaplanan Analiz Değerleri")
    
    if not edited_df.empty:
        totals_ww = {}
        for idx, row in edited_df.iterrows():
            h_ad = row["Hammadde Adı"]
            h_miktar = row["Miktar (% w/w)"]
            
            match = df_raw_db[df_raw_db["Ad"] == h_ad]
            if not match.empty:
                h_row = match.iloc[0]
                for col in df_raw_db.columns:
                    if col not in ["Kod", "Ad"]:
                        try:
                            val = float(h_row[col])
                        except (ValueError, TypeError):
                            val = 0.0
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

            zaman_kodu = datetime.datetime.now().strftime("%y%m%d-%H%M")
            sistem_kod = f"{tip}-{zaman_kodu}R0"

            st.session_state["current_calc_df"] = df_analysis_out
            st.session_state["current_meta"] = {
                "urun_adi": urun_adi,
                "sistem_kod": sistem_kod,
                "firma": firma,
                "yogunluk": yogunluk,
                "ph_degeri": ph_degeri,
                "talep_no": talep_no
            }

    if st.button("💾 Formülü Onayla ve Arşive Kaydet", use_container_width=True):
        # 1. Formülü veritabanına kaydet
        st.session_state["FORMULLER_DB"].append({
            "Kod": sistem_kod,
            "Ürün Adı": urun_adi,
            "Firma": firma,
            "Yoğunluk": yogunluk,
            "pH": ph_degeri,
            "Talep No": talep_no,
            "Tarih": datetime.datetime.now().strftime("%d.%m.%Y %H:%M")
        })
        
        # 2. TALEPLER_DB Durumunu Güncelle
        df_talepler = st.session_state["TALEPLER_DB"]
        idx_match = df_talepler[df_talepler["Talep No"] == talep_no].index
        if not idx_match.empty:
            st.session_state["TALEPLER_DB"].loc[idx_match[0], "Durum"] = "2 - Arşive Eklendi (Onay Bekliyor)"
            st.session_state["TALEPLER_DB"].loc[idx_match[0], "Üretilen Kod"] = sistem_kod

        st.balloons()
        st.success(f"Formülasyon Başarıyla Kaydedildi! | **Sistem Kodu:** `{sistem_kod}` | **Talep Durumu Güncellendi!**")

# ==========================================
# MODÜL 2: BELGE ÜRETİMİ
# ==========================================
elif secim == "📄 Belge Üretimi (CoA/CoC/SOP)":
    st.markdown("<h3 style='color:#10365C;'>📄 Otomatik Belge & Analiz Sertifikası Üretici</h3>", unsafe_allow_html=True)
    
    if "current_calc_df" not in st.session_state:
        st.info("⚠️ Henüz aktif bir formülasyon onaylanmadı. Lütfen önce **🧪 Formülasyon Tezgahı** sekmesinde bir reçete oluşturun.")
    else:
        meta = st.session_state["current_meta"]
        calc_df = st.session_state["current_calc_df"]
        
        st.success(f"Aktif Formülasyon: **{meta['urun_adi']}** | Kod: `{meta['sistem_kod']}` | Marka: **{meta['firma']}** | Talep: `{meta['talep_no']}`")
        
        doc_type = st.selectbox("Üretilecek Belge Türü", [
            "CoA - Certificate of Analysis (Analiz Sertifikası)",
            "CoC - Certificate of Conformity (Uygunluk Belgesi)",
            "SOP - Standard Operating Procedure (Laboratuvar/Üretim Talimatı)"
        ])
        
        st.warning("⚠️ **Geliştirme Notu:** Sayfa tasarımı ve kurumsal logo yerleşimi adım adım özelleştirilecektir.")

# ==========================================
# MODÜL 3: REÇETE VERİTABANI
# ==========================================
elif secim == "📚 Reçete Veritabanı (FORMULLER_DB)":
    st.markdown("<h3 style='color:#10365C;'>📚 Reçete Veritabanı (FORMULLER_DB)</h3>", unsafe_allow_html=True)
    db = st.session_state["FORMULLER_DB"]
    if not db:
        st.info("📂 Veritabanında kayıtlı formül bulunmuyor.")
    else:
        st.dataframe(pd.DataFrame(db), use_container_width=True)

# ==========================================
# MODÜL 4: HAMMADDE YÖNETİMİ
# ==========================================
elif secim == "⚙️ Hammadde & Analiz Veritabanı":
    st.markdown("<h3 style='color:#10365C;'>⚙️ Canlı Hammadde & Analiz Veritabanı Düzenleyici</h3>", unsafe_allow_html=True)
    updated_raw_db = st.data_editor(
        st.session_state["df_raw_db"],
        num_rows="dynamic",
        use_container_width=True,
        key="raw_db_editor"
    )
    if st.button("💾 Veritabanı Değişikliklerini Uygula & Kaydet", use_container_width=True):
        st.session_state["df_raw_db"] = updated_raw_db
        st.success("✅ Hammadde veritabanı başarıyla güncellendi!")
        st.rerun()

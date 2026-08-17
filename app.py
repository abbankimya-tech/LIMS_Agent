import streamlit as st
import pandas as pd
import datetime

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

# Giriş Kontrolü
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
# 2. HAMMADDELER KÜTÜPHANESİ (VERİTABANI)
# ==========================================
HAMMADDELER_LISTESI = [
    "Su (Deiyonize / Su)",
    "Enzimatik Protein Hidrolizatı",
    "Bitkisel Menşeli Amino Asit",
    "Çinko Sülfat Monohidrat",
    "Borik Asit",
    "Mono Potasyum Fosfat (MKP)",
    "Üre",
    "Potasyum Hidroksit (KOH)",
    "Sitrik Asit Anhidrit",
    "EDDHA-Fe (%6 o-o)",
    "Magnezyum Sülfat Heptahidrat",
    "Mangan Sülfat Monohidrat",
    "Humik Asit / Potasyum Humat",
    "Fosforik Asit (%85)",
    "Salisilik Asit",
    "Sodyum Lignosülfonat",
    "Surfaktan / Yayıcı Yapıştırıcı",
    "Diğer / Özel İnce Kimyasal"
]

# ==========================================
# 3. ANA UYGULAMA VE YAN MENÜ
# ==========================================
st.sidebar.title(f"👤 Kullanıcı: {st.session_state['username']}")
if st.sidebar.button("Güvenli Çıkış"):
    st.session_state["logged_in"] = False
    st.rerun()

st.sidebar.markdown("---")
secim = st.sidebar.radio("İşlem Modülü:", [
    "🧪 Formülasyon Tezgahı", 
    "📄 Belge Üretimi (CoA/CoC/SOP)", 
    "🤖 Gemini AI Asistanı"
])

# ==========================================
# MODÜL 1: FORMÜLASYON TEZGAHI & ANALİZ MOTORU
# ==========================================
if secim == "🧪 Formülasyon Tezgahı":
    st.markdown("<h3 style='color:#10365C;'>🧪 Formülasyon Oluşturma & Analiz Tezgahı</h3>", unsafe_allow_html=True)
    
    # 1. Ürün Başlık Verileri
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
    
    # Varsayılan Şablon Tablo
    if "df_recipe" not in st.session_state:
        st.session_state["df_recipe"] = pd.DataFrame([
            {"Hammadde Adı": "Su (Deiyonize / Su)", "Miktar (% w/w)": 45.00},
            {"Hammadde Adı": "Enzimatik Protein Hidrolizatı", "Miktar (% w/w)": 30.00},
            {"Hammadde Adı": "Çinko Sülfat Monohidrat", "Miktar (% w/w)": 25.00}
        ])

    # Canlı Düzenlenebilir Tablo (Açılır Liste Hammadde Seçimi İle)
    edited_df = st.data_editor(
        st.session_state["df_recipe"],
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "Hammadde Adı": st.column_config.SelectboxColumn(
                "Hammadde Adı",
                help="Reçeteye eklenecek hammaddeyi seçiniz",
                options=HAMMADDELER_LISTESI,
                required=True
            ),
            "Miktar (% w/w)": st.column_config.NumberColumn(
                "Miktar (% w/w)",
                help="Ağırlıkça yüzde oranı",
                min_value=0.0,
                max_value=100.0,
                step=0.1,
                format="%.2f %%"
            )
        }
    )
    
    # 100 kg Bakiye Kontrolü
    toplam_w_w = edited_df["Miktar (% w/w)"].sum() if not edited_df.empty else 0.0
    
    b1, b2 = st.columns([3, 1])
    with b1:
        if abs(toplam_w_w - 100.0) < 0.001:
            st.success(f"✅ TOPLAM BİLEŞİM: %{toplam_w_w:.2f} (w/w) — Reçete Tamamlandı!")
        elif toplam_w_w < 100.0:
            st.warning(f"⚠️ TOPLAM BİLEŞİM: %{toplam_w_w:.2f} (w/w) — Kalan Eksik: %{100.0 - toplam_w_w:.2f}")
        else:
            st.error(f"❌ TOPLAM BİLEŞİM: %{toplam_w_w:.2f} (w/w) — Fazla Miktar: %{toplam_w_w - 100.0:.2f}")

    # Otomatik Kod Üretimi
    zaman_kodu = datetime.datetime.now().strftime("%y%m%d-%H%M")
    sistem_kod = f"{tip}-{zaman_kodu}R0"
    revizyon_kod = "Rev-00"

    # --- CANLI HESAPLANAN ANALİZ TABLOSU ---
    st.markdown("---")
    st.write("##### 📊 Analiz Sonuç Tablosu (Otomatik Dönüştürülen Analiz Değerleri)")
    
    if not edited_df.empty:
        calc_df = edited_df.copy()
        calc_df["Miktar (% w/v)"] = calc_df["Miktar (% w/w)"] * yogunluk
        calc_df["Miktar (g/L)"] = calc_df["Miktar (% w/v)"] * 10

        st.dataframe(
            calc_df.style.format({
                "Miktar (% w/w)": "{:.2f} %",
                "Miktar (% w/v)": "{:.2f} %",
                "Miktar (g/L)": "{:.2f} g/L"
            }),
            use_container_width=True
        )

    if st.button("💾 Formülü Onayla ve Sistem Kodu Üret", use_container_width=True):
        st.balloons()
        st.success(f"Formülasyon Başarıyla Kaydedildi! | **Sistem Kodu:** `{sistem_kod}` | **Revizyon:** `{revizyon_kod}` | **Marka:** `{firma}`")

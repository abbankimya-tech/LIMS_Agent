import streamlit as st
import pandas as pd
import datetime
import io
from docx import Document
from docx.shared import Inches, Pt, RGBColor
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

# Veritabanı Hafızası
if "FORMULLER_DB" not in st.session_state:
    st.session_state["FORMULLER_DB"] = []

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
# 3. WORD BELGESİ OLUŞTURMA FONKSİYONU
# ==========================================
def generate_docx_document(doc_type, urun_adi, sistem_kod, firma, yogunluk, ph_degeri, df_data):
    doc = Document()
    
    title = doc.add_paragraph()
    title_run = title.add_run(f"{firma.upper()}\nLABORATORY QUALITY CONTROL & R&D DEPARTMENT")
    title_run.bold = True
    title_run.font.size = Pt(14)
    title_run.font.color.rgb = RGBColor(16, 54, 92)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    doc.add_paragraph("----------------------------------------------------------------------------------------------------")
    
    subtitle = doc.add_paragraph()
    sub_run = subtitle.add_run(f"DOCUMENT TYPE: {doc_type}")
    sub_run.bold = True
    sub_run.font.size = Pt(12)
    
    p_info = doc.add_paragraph()
    p_info.add_run(f"Product Name: {urun_adi}\n").bold = True
    p_info.add_run(f"System Code / Lot No: {sistem_kod}\n")
    p_info.add_run(f"Density (d): {yogunluk:.3f} g/cm³\n")
    p_info.add_run(f"pH (1/10): {ph_degeri:.1f}\n")
    p_info.add_run(f"Date: {datetime.datetime.now().strftime('%d.%m.%Y')}\n")
    
    doc.add_heading("Formulation & Specification Data", level=2)
    
    table = doc.add_table(rows=1, cols=4)
    table.style = 'Table Grid'
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = 'Component / Raw Material'
    hdr_cells[1].text = 'w/w %'
    hdr_cells[2].text = 'w/v %'
    hdr_cells[3].text = 'g/L'
    
    for row in table.rows[0].cells:
        for p in row.paragraphs:
            for run in p.runs:
                run.font.bold = True
                
    for idx, row in df_data.iterrows():
        row_cells = table.add_row().cells
        row_cells[0].text = str(row["Hammadde Adı"])
        row_cells[1].text = f"{row['Miktar (% w/w)']:.2f}"
        row_cells[2].text = f"{row['Miktar (% w/v)']:.2f}"
        row_cells[3].text = f"{row['Miktar (g/L)']:.2f}"
        
    doc.add_paragraph("\nApproved by Head of Quality Control & R&D Laboratory")
    
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# ==========================================
# 4. ANA UYGULAMA VE YAN MENÜ
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
    "🤖 Gemini AI Ar-Ge Asistanı"
])

# ==========================================
# MODÜL 1: FORMÜLASYON TEZGAHI & ANALİZ MOTORU
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
            {"Hammadde Adı": "Su (Deiyonize / Su)", "Miktar (% w/w)": 45.00},
            {"Hammadde Adı": "Enzimatik Protein Hidrolizatı", "Miktar (% w/w)": 30.00},
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
    
    b1, b2 = st.columns([3, 1])
    with b1:
        if abs(toplam_w_w - 100.0) < 0.001:
            st.success(f"✅ TOPLAM BİLEŞİM: %{toplam_w_w:.2f} (w/w) — Reçete Tamamlandı!")
        elif toplam_w_w < 100.0:
            st.warning(f"⚠️ TOPLAM BİLEŞİM: %{toplam_w_w:.2f} (w/w) — Kalan Eksik: %{100.0 - toplam_w_w:.2f}")
        else:
            st.error(f"❌ TOPLAM BİLEŞİM: %{toplam_w_w:.2f} (w/w) — Fazla Miktar: %{toplam_w_w - 100.0:.2f}")

    zaman_kodu = datetime.datetime.now().strftime("%y%m%d-%H%M")
    sistem_kod = f"{tip}-{zaman_kodu}R0"

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

        st.session_state["current_calc_df"] = calc_df
        st.session_state["current_meta"] = {
            "urun_adi": urun_adi,
            "sistem_kod": sistem_kod,
            "firma": firma,
            "yogunluk": yogunluk,
            "ph_degeri": ph_degeri,
            "talep_no": talep_no
        }

    if st.button("💾 Formülü Onayla ve Veritabanına Kaydet", use_container_width=True):
        st.session_state["FORMULLER_DB"].append({
            "Kod": sistem_kod,
            "Ürün Adı": urun_adi,
            "Firma": firma,
            "Yoğunluk": yogunluk,
            "pH": ph_degeri,
            "Tarih": datetime.datetime.now().strftime("%d.%m.%Y %H:%M"),
            "Detay": calc_df
        })
        st.balloons()
        st.success(f"Formülasyon Başarıyla `FORMULLER_DB` Veritabanına Kaydedildi! | **Kod:** `{sistem_kod}`")

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
        
        st.success(f"Aktif Formülasyon: **{meta['urun_adi']}** | Kod: `{meta['sistem_kod']}` | Marka: **{meta['firma']}**")
        
        doc_type = st.selectbox("Üretilecek Belge Türü", [
            "CoA - Certificate of Analysis (Analiz Sertifikası)",
            "CoC - Certificate of Conformity (Uygunluk Belgesi)",
            "SOP - Standard Operating Procedure (Laboratuvar/Üretim Talimatı)"
        ])
        
        if st.button("🚀 Word Belgesini Derle ve İndir", use_container_width=True):
            word_file = generate_docx_document(
                doc_type.split(" - ")[0],
                meta["urun_adi"],
                meta["sistem_kod"],
                meta["firma"],
                meta["yogunluk"],
                meta["ph_degeri"],
                calc_df
            )
            
            file_name = f"{meta['firma'].split()[0]}_{doc_type.split(' - ')[0]}_{meta['urun_adi']}.docx"
            
            st.download_button(
                label=f"📥 {file_name} Dosyasını Bilgisayara İndir",
                data=word_file,
                file_name=file_name,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )

# ==========================================
# MODÜL 3: REÇETE VERİTABANI
# ==========================================
elif secim == "📚 Reçete Veritabanı (FORMULLER_DB)":
    st.markdown("<h3 style='color:#10365C;'>📚 Reçete Veritabanı & Arşiv Takibi</h3>", unsafe_allow_html=True)
    
    db = st.session_state["FORMULLER_DB"]
    
    if not db:
        st.info("📂 Veritabanında henüz kayıtlı formül bulunmuyor. Formülasyon Tezgahı'nda reçete oluşturup kaydedebilirsiniz.")
    else:
        st.write(f"**Kayıtlı Toplam Formülasyon Sayısı:** {len(db)}")
        
        summary_data = []
        for item in db:
            summary_data.append({
                "Sistem Kodu": item["Kod"],
                "Ürün Adı": item["Ürün Adı"],
                "Firma": item["Firma"],
                "Yoğunluk (d)": item["Yoğunluk"],
                "pH": item["pH"],
                "Kayıt Tarihi": item["Tarih"]
            })
            
        df_summary = pd.DataFrame(summary_data)
        st.dataframe(df_summary, use_container_width=True)
        
        st.markdown("---")
        secilen_kod = st.selectbox("Detayını İncelemek İstediğiniz Formülü Seçin:", df_summary["Sistem Kodu"])
        
        for item in db:
            if item["Kod"] == secilen_kod:
                st.write(f"##### 🔎 Reçete Detayı: {item['Ürün Adı']} (`{item['Kod']}`)")
                st.dataframe(item["Detay"], use_container_width=True)

# ==========================================
# MODÜL 4: GEMINI AI ASİSTANI
# ==========================================
elif secim == "🤖 Gemini AI Ar-Ge Asistanı":
    st.markdown("<h3 style='color:#10365C;'>🤖 Gemini AI Ar-Ge & Kimya Asistanı</h3>", unsafe_allow_html=True)
    
    st.write("Formülasyonlarınızın kimyasal uyumluluğu, hammadde ikameleri ve dozaj optimizasyonları için akıllı danışmanınız.")
    
    user_query = st.text_area("Ar-Ge / Kimya Sorunuzu Giriniz:", "Örn: Çinko sülfat ile amino asit şelatlama reaksiyonunda pH dengesini korumak için hangi tampon çözelti önerilir?")
    
    if st.button("🧠 Yapay Zekaya Danış", use_container_width=True):
        st.info("💡 **Ar-Ge Danışmanı Yanıtı:**\n\nÇinko sülfat monohidrat ile amino asit kompleksleşmesinde pH 5.5 - 6.5 aralığı ideal stabiliteler sunar. Çökeltiyi önlemek için ortama %1-2 oranında Sitrik Asit eklenmesi şelat yapısını destekler.")

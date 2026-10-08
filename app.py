import streamlit as st
import datetime
import io
import os
import re
from PIL import Image
from pypdf import PdfReader
import pytesseract
from builder import generate_pdf_contract, generate_docx_contract, PRESTATORI_DATA

st.set_page_config(
    page_title="Generator Contracte Contabilitate",
    page_icon="💼",
    layout="wide"
)

# Custom Styling cu font Helvetica
st.markdown("""
<style>
    * {
        font-family: 'Helvetica', 'Arial', sans-serif !important;
    }
    .main-header {
        background: linear-gradient(135deg, #1B365D 0%, #2B4C7E 100%);
        padding: 22px 28px;
        border-radius: 12px;
        color: white;
        margin-bottom: 25px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.07);
    }
    .main-header h1 {
        color: #FFFFFF !important;
        margin: 0;
        font-size: 26px;
        font-weight: 700;
    }
    .main-header p {
        color: #E2E8F0;
        margin: 6px 0 0 0;
        font-size: 14px;
    }
    .upload-box {
        background-color: #F8FAFC;
        border: 2px dashed #CBD5E1;
        padding: 16px;
        border-radius: 8px;
        margin-bottom: 18px;
    }
</style>
<div class="main-header">
    <h1>💼 Generator Automat Contracte Servicii Contabilitate</h1>
    <p>Aplicație colaborativă pentru birouri contabile • Încărcare copie CUI & Completare Asistată • Export PDF & Word (Font Helvetica)</p>
</div>
""", unsafe_allow_html=True)

def parse_cui_text(raw_text):
    """Extrage automat datele de identificare ale firmei dintr-un text de Certificat de Înregistrare (CUI)."""
    data = {}
    lines = [line.strip() for line in raw_text.splitlines() if line.strip()]

    # 1. Denumire / Firmă
    for line in lines:
        m = re.search(r'^(?:Firmă|Firma|Denumire|Societatea)\s*:\s*(.+)$', line, re.IGNORECASE)
        if m:
            data['nume'] = m.group(1).strip()
            break
    if 'nume' not in data:
        for line in lines:
            m_srl = re.search(r'([A-Z0-9\s\.\-]{3,60}(?:S\.?R\.?L\.?|S\.?A\.?))', line, re.IGNORECASE)
            if m_srl and not any(k in line.lower() for k in ['certificat', 'romania', 'ministerul', 'registrul', 'justitiei']):
                data['nume'] = m_srl.group(1).strip()
                break

    # 2. CUI / Cod Unic de Înregistrare
    m_cui = re.search(r'(?:Cod\s*Unic\s*de\s*Înregistrare|Cod\s*Unic\s*de\s*Inregistrare|C\.?U\.?I\.?|C\.?I\.?F\.?)\s*:\s*(?:RO)?\s*([0-9]{6,10})', raw_text, re.IGNORECASE)
    if m_cui:
        data['cui'] = m_cui.group(1).strip()
    else:
        m_cui_fall = re.search(r'\b(?:RO)?([0-9]{6,10})\b', raw_text)
        if m_cui_fall:
            data['cui'] = m_cui_fall.group(1).strip()

    # 3. Nr. Registrul Comerțului
    m_rc = re.search(r'(?:Nr\.?\s*de\s*ordine\s*în\s*registrul\s*comerțului|Nr\.?\s*de\s*ordine\s*in\s*registrul\s*comertului|Reg\.?\s*Com\.?)\s*:\s*([A-Z0-9\/\.\-]+)', raw_text, re.IGNORECASE)
    if m_rc:
        data['reg_com'] = m_rc.group(1).strip()
    else:
        m_rc_j = re.search(r'\b([JF]\d{1,2}\/\d+\/\d{4})\b', raw_text)
        if m_rc_j:
            data['reg_com'] = m_rc_j.group(1).strip()
        else:
            m_rc_long = re.search(r'\b(J\d{12,14})\b', raw_text)
            if m_rc_long:
                data['reg_com'] = m_rc_long.group(1).strip()

    # 4. Sediu social
    m_sediu = re.search(r'Sediu\s*social\s*:\s*(.+?)(?=(?:\n\s*(?:Activitatea|Cod Unic|Identificator|Oficiul|Nr\. de ordine|Data eliberării|Seria))|\Z)', raw_text, re.IGNORECASE | re.DOTALL)
    if m_sediu:
        data['sediu'] = ' '.join(m_sediu.group(1).split()).strip()

    return data

def extract_text_from_upload(uploaded_file):
    """Extrage textul dintr-un fișier PDF sau Imagine încărcată de utilizator."""
    file_bytes = uploaded_file.read()
    fname = uploaded_file.name.lower()
    raw_text = ""

    if fname.endswith(".pdf"):
        # 1. Încercare extragere text nativ (digital PDF)
        try:
            reader = PdfReader(io.BytesIO(file_bytes))
            for page in reader.pages:
                t = page.extract_text()
                if t:
                    raw_text += t + "\n"
        except Exception:
            pass

        # 2. Dacă PDF-ul este scanat (fără text nativ), încercare OCR
        if not raw_text.strip():
            try:
                import pdfplumber
                with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
                    for page in pdf.pages[:2]:  # Primele 2 pagini
                        img = page.to_image(resolution=200).original
                        raw_text += pytesseract.image_to_string(img) + "\n"
            except Exception:
                pass

    elif any(fname.endswith(ext) for ext in [".png", ".jpg", ".jpeg", ".webp", ".bmp", ".tiff"]):
        try:
            img = Image.open(io.BytesIO(file_bytes))
            raw_text = pytesseract.image_to_string(img)
        except Exception:
            pass

    return raw_text

# Inițializare Session State pentru datele clientului
if "client_data" not in st.session_state:
    st.session_state["client_data"] = {
        "nume": "",
        "sediu": "",
        "reg_com": "",
        "cui": "",
        "iban": "",
        "banca": "",
        "reprezentant": "",
        "functie": "Administrator",
        "telefon": "",
        "email": ""
    }

if "last_uploaded_file" not in st.session_state:
    st.session_state["last_uploaded_file"] = None

col_prest, col_client = st.columns([1, 1], gap="large")

with col_prest:
    st.subheader("🏢 1. Selectare Prestator & Setări Contract")
    prest_nume = st.selectbox(
        "Alegeți societatea prestatoare:",
        list(PRESTATORI_DATA.keys()),
        index=0
    )
    prest = PRESTATORI_DATA[prest_nume]

    # Previzualizare sigle
    c_logo1, c_logo2 = st.columns([2, 1])
    with c_logo1:
        if os.path.exists(prest["logo_path"]):
            st.image(prest["logo_path"], caption=f"Siglă {prest['nume']}", width=220)
    with c_logo2:
        ceccar_p = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "logo_ceccar.png")
        if os.path.exists(ceccar_p):
            st.image(ceccar_p, caption="Siglă CECCAR", width=70)

    st.markdown("---")
    st.markdown("##### ⚙️ Parametri Contractuali")
    c_p1, c_p2 = st.columns(2)
    with c_p1:
        nr_contract = st.text_input("Număr Contract", value=f"2026-{datetime.date.today().strftime('%m%d')}")
        pret_baza = st.number_input("Preț Lunar de Bază (LEI)", value=500, step=50, min_value=100)
    with c_p2:
        data_contract = st.date_input("Data Încheierii", value=datetime.date.today()).strftime("%d.%m.%Y")
        trans_incluse = st.selectbox("Tranșă tranzacții de bază", ["0 - 40", "40 - 100", "101 - 200"], index=0)

    data_start = st.text_input("Data Început Prestații (opțional)", value=data_contract)

with col_client:
    st.subheader("📄 2. Date Beneficiar (Client)")
    st.markdown("Puteți încărca copia Certificatului de Înregistrare (CUI) pentru completare automată sau puteți tasta datele direct manual:")

    uploaded_cui = st.file_uploader(
        "📎 Încarcă copie CUI (PDF, Scan sau Imagine PNG/JPG):",
        type=["pdf", "png", "jpg", "jpeg", "webp"],
        key="cui_uploader",
        help="Aplicația extrage automat Denumirea firmei, CUI-ul, Numărul de Registru al Comerțului și Sediul Social."
    )

    if uploaded_cui is not None and uploaded_cui.name != st.session_state["last_uploaded_file"]:
        with st.spinner("Se procesează copia CUI și se extrag datele firmei..."):
            extracted_text = extract_text_from_upload(uploaded_cui)
            if extracted_text.strip():
                parsed = parse_cui_text(extracted_text)
                if parsed:
                    if parsed.get("nume"):
                        st.session_state["client_data"]["nume"] = parsed["nume"]
                    if parsed.get("cui"):
                        st.session_state["client_data"]["cui"] = parsed["cui"]
                    if parsed.get("reg_com"):
                        st.session_state["client_data"]["reg_com"] = parsed["reg_com"]
                    if parsed.get("sediu"):
                        st.session_state["client_data"]["sediu"] = parsed["sediu"]
                    st.session_state["last_uploaded_file"] = uploaded_cui.name
                    st.success(f"✅ Date extrase cu succes din fișierul încărcat: **{parsed.get('nume', 'Client')}**")
                else:
                    st.warning("Fișierul a fost citit, dar unele câmpuri nu au putut fi identificate automat. Vă rugăm să le completați manual mai jos.")
            else:
                st.warning("Nu s-a putut extrage text automat din acest fișier. Vă rugăm să completați datele manual în casetele de mai jos.")

    st.markdown("##### 🏢 Date Firmă Client (Editabile manual)")
    cl_nume = st.text_input("Denumire Societate Client", value=st.session_state["client_data"]["nume"], placeholder="ex: ARBITRANS BUSINESS SRL")
    c_id1, c_id2 = st.columns(2)
    with c_id1:
        cl_cui = st.text_input("Cod Unic de Înregistrare (CUI)", value=st.session_state["client_data"]["cui"], placeholder="ex: 15771194")
    with c_id2:
        cl_regcom = st.text_input("Nr. Registrul Comerțului", value=st.session_state["client_data"]["reg_com"], placeholder="ex: J2003012923400 sau J40/12923/2003")

    cl_sediu = st.text_input("Adresă Sediu Social Completă", value=st.session_state["client_data"]["sediu"], placeholder="ex: Bucureşti Sectorul 1, Str. Băiculeşti Nr. 7...")

    st.markdown("##### 👤 Reprezentant & Date Bancare / Contact")
    c_adm1, c_adm2 = st.columns(2)
    with c_adm1:
        cl_adm = st.text_input("Nume & Prenume Administrator", value=st.session_state["client_data"]["reprezentant"], placeholder="ex: Popescu Ion")
    with c_adm2:
        cl_fnc = st.text_input("Funcție", value=st.session_state["client_data"]["functie"])

    c_bk1, c_bk2 = st.columns(2)
    with c_bk1:
        cl_banca = st.text_input("Banca", value=st.session_state["client_data"]["banca"], placeholder="ex: Banca Transilvania")
    with c_bk2:
        cl_iban = st.text_input("Cont IBAN", value=st.session_state["client_data"]["iban"], placeholder="RO__BTRL...")

    c_ct1, c_ct2 = st.columns(2)
    with c_ct1:
        cl_tel = st.text_input("Telefon Contact", value=st.session_state["client_data"]["telefon"], placeholder="ex: 0722...")
    with c_ct2:
        cl_mail = st.text_input("E-mail Oficial Comunicare", value=st.session_state["client_data"]["email"], placeholder="client@firma.ro")

final_client = {
    "nume": cl_nume or "____________________ S.R.L.",
    "sediu": cl_sediu or "____________________",
    "reg_com": cl_regcom or "J___/______/________",
    "cui": cl_cui or "________",
    "iban": cl_iban or "RO____________________",
    "banca": cl_banca or "____________________",
    "reprezentant": cl_adm or "____________________",
    "functie": cl_fnc or "Administrator",
    "telefon": cl_tel or "__________",
    "email": cl_mail or "____________________"
}

st.markdown("---")
st.subheader("📥 3. Generare & Descărcare Contract")
st.markdown("Contractul generat include toate clauzele din documentul cadru final, fontul **Helvetica**, siglele oficiale în antet, clauza critică de 23 ale lunii, penalitățile moderate de 15% și 30%, precum și cele trei Anexe oficiale 2026.")

c_btn1, c_btn2, _ = st.columns([1.5, 1.5, 2])

# Nume fișiere curate
clean_name = re.sub(r'[^a-zA-Z0-9]', '_', final_client['nume'])[:15]
nume_pdf = f"Contract_{prest['nume'].split()[0]}_{clean_name}_{nr_contract}.pdf"
nume_docx = f"Contract_{prest['nume'].split()[0]}_{clean_name}_{nr_contract}.docx"

pdf_bytes = generate_pdf_contract(
    prestator_key=prest_nume,
    client_data=final_client,
    nr_ctr=nr_contract,
    data_ctr=data_contract,
    pret_baza=pret_baza,
    trans_incluse=trans_incluse,
    data_start=data_start
)

docx_bytes = generate_docx_contract(
    prestator_key=prest_nume,
    client_data=final_client,
    nr_ctr=nr_contract,
    data_ctr=data_contract,
    pret_baza=pret_baza,
    trans_incluse=trans_incluse,
    data_start=data_start
)

with c_btn1:
    st.download_button(
        label="📄 Descarcă Contract PDF Editabil (Helvetica)",
        data=pdf_bytes,
        file_name=nume_pdf,
        mime="application/pdf",
        use_container_width=True
    )

with c_btn2:
    st.download_button(
        label="📝 Descarcă Contract Word (.docx)",
        data=docx_bytes,
        file_name=nume_docx,
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True
    )

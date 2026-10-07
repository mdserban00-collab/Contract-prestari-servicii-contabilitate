import streamlit as st
import requests
import datetime
import io
import os
import re
from builder import generate_pdf_contract, generate_docx_contract, PRESTATORI_DATA

st.set_page_config(
    page_title="Generator Contracte Contabilitate",
    page_icon="💼",
    layout="wide"
)

# Custom Styling
st.markdown("""
<style>
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
    .risk-box-danger {
        background-color: #FEF2F2;
        border-left: 5px solid #DC2626;
        padding: 14px 18px;
        border-radius: 6px;
        margin-bottom: 15px;
    }
    .risk-box-success {
        background-color: #F0FDF4;
        border-left: 5px solid #16A34A;
        padding: 14px 18px;
        border-radius: 6px;
        margin-bottom: 15px;
    }
</style>
<div class="main-header">
    <h1>💼 Generator Automat Contracte Servicii Contabilitate</h1>
    <p>Aplicație colaborativă securizată pentru birouri contabile • Interogare automată CUI la ANAF • Rubrica Risc & Conformitate • Export PDF & Word</p>
</div>
""", unsafe_allow_html=True)

def analyze_anaf_risks(item):
    riscuri = []
    detalii = {}

    # 1. Date generale & stare inregistrare
    dg = item.get("date_generale", {})
    stare_inreg = (dg.get("stare_inregistrare") or item.get("stare_inregistrare") or "").upper()

    # 2. Inactivitate fiscala
    st_inactiv = item.get("stare_inactiv", {})
    status_inactiv = st_inactiv.get("statusInactivi") or item.get("statusInactivi") or False
    data_inactivare = st_inactiv.get("dataInactivare") or ""
    data_reactivare = st_inactiv.get("dataReactivare") or ""

    if status_inactiv or "INACTIV" in stare_inreg:
        msg = "Firma figurează INACTIVĂ FISCAL în Registrul ANAF"
        if data_inactivare:
            msg += f" din data de {data_inactivare}"
        riscuri.append(("INACTIVITATE FISCALĂ", msg, "CRITIC"))
        detalii["inactiv"] = True
    else:
        detalii["inactiv"] = False

    # 3. Insolventa
    st_insolv = item.get("stare_insolventa", {})
    status_insolv = st_insolv.get("statusInsolventa") or item.get("statusInsolventa") or False
    data_insolv = st_insolv.get("dataInsolventa") or ""

    if status_insolv or "INSOLVENT" in stare_inreg:
        msg = "Firma se află în procedură judiciară de INSOLVENȚĂ"
        if data_insolv:
            msg += f" din data de {data_insolv}"
        riscuri.append(("INSOLVENȚĂ", msg, "CRITIC"))
        detalii["insolventa"] = True
    else:
        detalii["insolventa"] = False

    # 4. Radiere / Lichidare / Dizolvare / Faliment
    st_rad = item.get("stare_radiere", {})
    status_rad = st_rad.get("statusRadiere") or item.get("statusRadiere") or False
    data_rad = st_rad.get("dataRadiere") or st_inactiv.get("dataRadiere") or ""

    cuvinte_lichidare = ["RADIAT", "LICHIDARE", "FALIMENT", "DIZOLVARE"]
    este_lichidare = status_rad or any(w in stare_inreg for w in cuvinte_lichidare)
    if este_lichidare:
        tip = "RADIATĂ / ÎN LICHIDARE / FALIMENT"
        for w in cuvinte_lichidare:
            if w in stare_inreg:
                tip = f"ÎN STARE DE {w}"
                break
        msg = f"Firma figurează {tip}"
        if data_rad:
            msg += f" din data de {data_rad}"
        riscuri.append(("LICHIDARE / RADIERE", msg, "CRITIC"))
        detalii["lichidare"] = True
    else:
        detalii["lichidare"] = False

    # 5. Cod TVA anulat din oficiu
    st_tva = item.get("inregistrare_scop_Tva", {})
    platitor_tva = st_tva.get("scpTVA", False)
    perioade_tva = st_tva.get("perioade_tva", []) or []

    tva_anulat = False
    data_anulare_tva = ""
    mesaj_anulare = ""

    if not platitor_tva and perioade_tva:
        for p in perioade_tva:
            d_anul = p.get("data_anul_imp_ScpTVA") or p.get("data_sfarsit_ScpTVA")
            m_scop = p.get("mesaj_scop_tva") or ""
            if d_anul or "ANUL" in m_scop.upper() or "316" in m_scop:
                tva_anulat = True
                data_anulare_tva = d_anul or ""
                mesaj_anulare = m_scop
                break

    if tva_anulat:
        msg = "Codul de TVA a fost ANULAT DIN OFICIU de către organele fiscale"
        if data_anulare_tva:
            msg += f" la data de {data_anulare_tva}"
        if mesaj_anulare:
            msg += f" ({mesaj_anulare.strip()})"
        riscuri.append(("COD TVA ANULAT DIN OFICIU", msg, "CRITIC"))
        detalii["tva_anulat"] = True
    else:
        detalii["tva_anulat"] = False

    detalii["platitor_tva"] = platitor_tva
    detalii["stare_inregistrare"] = stare_inreg

    # Generare text sintetic pentru contract
    if riscuri:
        sumare = [r[1] for r in riscuri]
        risc_text = "ATENȚIE - Mențiuni speciale: " + "; ".join(sumare)
    else:
        risc_text = "Activă fiscal conform ANAF, fără proceduri de insolvență/lichidare sau cod TVA anulat."

    return riscuri, detalii, risc_text

def fetch_anaf_company(cui_input):
    clean_cui = re.sub(r'\D', '', str(cui_input))
    if not clean_cui:
        return None, "Introduceți un CUI valid compus doar din cifre."
    url = "https://api.anaf.ro/PlatitorTvaRest/api/v8/ws/tva"
    payload = [{"cui": int(clean_cui), "data": datetime.date.today().strftime("%Y-%m-%d")}]
    try:
        resp = requests.post(url, json=payload, timeout=8)
        if resp.status_code == 200:
            data = resp.json()
            if "found" in data and len(data["found"]) > 0:
                item = data["found"][0]
                dg = item.get("date_generale", {})
                adresa_str = dg.get("adresa") or item.get("adresa") or ""
                nume_str = dg.get("denumire") or item.get("denumire") or ""
                regcom_str = dg.get("nrRegCom") or item.get("nrRegCom") or ""
                
                riscuri, detalii, risc_text = analyze_anaf_risks(item)

                return {
                    "nume": nume_str.strip(),
                    "sediu": adresa_str.strip(),
                    "reg_com": regcom_str.strip(),
                    "cui": clean_cui,
                    "riscuri": riscuri,
                    "detalii": detalii,
                    "risc_text": risc_text
                }, None
            return None, f"CUI-ul {clean_cui} nu a fost găsit în baza oficială ANAF."
        return None, f"Serverul ANAF a răspuns cu codul HTTP {resp.status_code}."
    except Exception as e:
        return None, f"Eroare de conexiune la serverul ANAF ({str(e)}). Vă rugăm să completați manual datele."

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
        "email": "",
        "risc": "Activă fiscal conform ANAF, fără proceduri de insolvență/lichidare sau cod TVA anulat."
    }

if "riscuri_gasite" not in st.session_state:
    st.session_state["riscuri_gasite"] = []

if "cui_verificat" not in st.session_state:
    st.session_state["cui_verificat"] = False

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
    st.subheader("🔍 2. Date Beneficiar (Client) & Verificare Risc")
    st.markdown("Introduceți CUI-ul clientului pentru preluarea automată a datelor și analiza riscului fiscal de la ANAF:")

    cc1, cc2 = st.columns([2.5, 1.5])
    with cc1:
        cui_input = st.text_input("CUI Client:", placeholder="ex: 34819582", key="cui_input_box")
    with cc2:
        st.write("")
        st.write("")
        btn_anaf = st.button("⚡ Preluare ANAF", use_container_width=True)

    if btn_anaf and cui_input:
        with st.spinner("Se interoghează baza de date ANAF și se analizează istoricul fiscal..."):
            res, err = fetch_anaf_company(cui_input)
            if res:
                st.session_state["cui_verificat"] = True
                st.session_state["client_data"]["nume"] = res["nume"]
                st.session_state["client_data"]["sediu"] = res["sediu"]
                st.session_state["client_data"]["cui"] = res["cui"]
                st.session_state["client_data"]["risc"] = res["risc_text"]
                st.session_state["riscuri_gasite"] = res["riscuri"]
                if res["reg_com"]:
                    st.session_state["client_data"]["reg_com"] = res["reg_com"]
                st.success(f"✅ Găsit la ANAF: **{res['nume']}**")
            else:
                st.session_state["cui_verificat"] = False
                st.error(f"⚠️ {err}")

    # ==================== RUBRICA RISC ====================
    if st.session_state["cui_verificat"]:
        if st.session_state["riscuri_gasite"]:
            st.error("🚨 **ALERTE DE RISC IDENTIFICATE LA ANAF PENTRU ACEST CLIENT:**")
            for titlu_r, desc_r, _ in st.session_state["riscuri_gasite"]:
                st.markdown(f"- 🛑 **{titlu_r}:** {desc_r}")
            st.warning("⚠️ **Atenționare Birou Contabilitate:** În caz de insolvență/lichidare, verificați dacă reprezentantul are drept de administrare sau dacă actele trebuie semnate de administratorul judiciar/lichidator. Pentru firme inactive sau cu TVA anulat se aplică regimul art. 11 Cod Fiscal (lipsă deductibilitate, obligație D311 etc.). Mențiunile au fost preluate automat în contract.")
        else:
            st.success("🟢 **RUBRICA RISC: RISC SCĂZUT CONFORM ANAF**\n\nFirma este **ACTIVĂ fiscal**, nu figurează în procedură de insolvență, dizolvare sau lichidare și nu are codul de TVA anulat din oficiu.")

    cl_nume = st.text_input("Denumire Societate Client", value=st.session_state["client_data"]["nume"])
    c_id1, c_id2 = st.columns(2)
    with c_id1:
        cl_cui = st.text_input("Cod Unic de Înregistrare (CUI)", value=st.session_state["client_data"]["cui"])
    with c_id2:
        cl_regcom = st.text_input("Nr. Registrul Comerțului", value=st.session_state["client_data"]["reg_com"])

    cl_sediu = st.text_input("Adresă Sediu Social Completă", value=st.session_state["client_data"]["sediu"])

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

    st.markdown("##### ⚠️ Mențiune Rubrica Risc (apare în contract)")
    cl_risc = st.text_area(
        "Stare fiscală & risc consemnată la data încheierii:",
        value=st.session_state["client_data"].get("risc", "Activă fiscal conform ANAF, fără proceduri de insolvență/lichidare sau cod TVA anulat."),
        help="Acest text este tipărit în Capitolul I al contractului la datele de identificare ale clientului.",
        height=70
    )

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
    "email": cl_mail or "____________________",
    "risc": cl_risc or "Activă fiscal conform ANAF, fără proceduri de insolvență/lichidare sau cod TVA anulat."
}

st.markdown("---")
st.subheader("📥 3. Generare & Descărcare Contract")
st.markdown("Contractul generat include toate clauzele contractuale din documentul final, clauza critică de 23 ale lunii, penalitățile moderate de 15% și 30%, Rubrica Risc a clientului, precum și cele trei Anexe oficiale 2026.")

c_btn1, c_btn2, _ = st.columns([1.5, 1.5, 2])

# Generate Files
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
        label="📄 Descarcă Contract PDF Editabil",
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

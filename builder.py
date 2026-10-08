import os
import io
import datetime
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml, OxmlElement

from reportlab.lib.pagesizes import A4
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas

from reportlab.pdfbase.pdfmetrics import registerFontFamily

# Register Helvetica TrueType fonts with full Romanian diacritics support
_assets_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'assets')
_h_reg = os.path.join(_assets_dir, 'Helvetica.ttf')
_h_bold = os.path.join(_assets_dir, 'Helvetica-Bold.ttf')
_h_it = os.path.join(_assets_dir, 'Helvetica-Oblique.ttf')
_h_bi = os.path.join(_assets_dir, 'Helvetica-BoldOblique.ttf')

# Fallbacks if assets are not local
if not os.path.exists(_h_reg):
    for _p in ['/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf', '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf']:
        if os.path.exists(_p): _h_reg = _p; break
if not os.path.exists(_h_bold):
    for _p in ['/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf', '/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf']:
        if os.path.exists(_p): _h_bold = _p; break
if not os.path.exists(_h_it):
    for _p in ['/usr/share/fonts/truetype/liberation/LiberationSans-Italic.ttf']:
        if os.path.exists(_p): _h_it = _p; break
if not os.path.exists(_h_bi):
    for _p in ['/usr/share/fonts/truetype/liberation/LiberationSans-BoldItalic.ttf']:
        if os.path.exists(_p): _h_bi = _p; break

pdfmetrics.registerFont(TTFont('Helvetica', _h_reg))
pdfmetrics.registerFont(TTFont('Helvetica-Bold', _h_bold))
if os.path.exists(_h_it):
    pdfmetrics.registerFont(TTFont('Helvetica-Oblique', _h_it))
if os.path.exists(_h_bi):
    pdfmetrics.registerFont(TTFont('Helvetica-BoldOblique', _h_bi))

registerFontFamily('Helvetica', 
                   normal='Helvetica', 
                   bold='Helvetica-Bold', 
                   italic='Helvetica-Oblique' if os.path.exists(_h_it) else 'Helvetica', 
                   boldItalic='Helvetica-BoldOblique' if os.path.exists(_h_bi) else 'Helvetica-Bold')

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ASSETS_DIR = os.path.join(BASE_DIR, "assets")

PRESTATORI_DATA = {
    "RIGHT WAY - TAX CONSULTING & MORE SRL": {
        "nume": "RIGHT WAY - TAX CONSULTING & MORE SRL",
        "sediu": "Corbeanca, Str. Iepurașului Nr. 9, Jud. Ilfov",
        "reg_com": "J2017005176238",
        "cui": "38351629",
        "ceccar": "Carnet / autorizație CECCAR seria A nr. 0013367",
        "iban": "RO93BTRLRONCRT0436157901",
        "banca": "Banca Transilvania",
        "reprezentant": "Trașcă - Șerban Maria - Daniela",
        "functie": "Administrator",
        "email": "taxe.conta@gmail.com",
        "web": "www.rightwaytax.ro",
        "telefon": "0722.522.167",
        "logo_path": os.path.join(ASSETS_DIR, "logo_rightway.png")
    },
    "REGNSKAB SRL": {
        "nume": "REGNSKAB SRL",
        "sediu": "București, Sos. Sălaj nr. 5, bl. 58A, sc. 1, et. 4, ap. 13, Sector 5",
        "reg_com": "J2014013701407",
        "cui": "33836240",
        "ceccar": "Carnet / autorizație CECCAR seria A nr. 0012462",
        "iban": "RO09BACX0000001082606000",
        "banca": "UniCredit Bank",
        "reprezentant": "Gudană Ștefania - Alina",
        "functie": "Administrator",
        "email": "regnskabsrl@gmail.com",
        "web": "www.contabsolution.ro",
        "telefon": "0741.122.219",
        "logo_path": os.path.join(ASSETS_DIR, "logo_regnskab.png")
    }
}

class NumberedCanvasWithLogos(canvas.Canvas):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._saved_page_states = []

    def showPage(self):
        self._saved_page_states.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        num_pages = len(self._saved_page_states)
        for state in self._saved_page_states:
            self.__dict__.update(state)
            self.draw_header_footer(num_pages)
            super().showPage()
        super().save()

    def draw_header_footer(self, page_count):
        self.saveState()
        w, h = A4
        prestator_logo = getattr(self, 'prestator_logo', '/workspace/scratch/assets/logo_rightway.png')
        ceccar_logo = os.path.join(ASSETS_DIR, 'logo_ceccar.png')

        # Running Header
        if os.path.exists(prestator_logo):
            self.drawImage(prestator_logo, 36, h - 48, width=145, height=33, preserveAspectRatio=True, mask='auto')
        if os.path.exists(ceccar_logo):
            self.drawImage(ceccar_logo, w - 36 - 44, h - 50, width=44, height=44, preserveAspectRatio=True, mask='auto')

        self.setStrokeColor(colors.HexColor('#CBD5E1'))
        self.setLineWidth(0.8)
        self.line(36, h - 54, w - 36, h - 54)

        # Running Footer
        self.line(36, 38, w - 36, 38)
        self.setFont('Helvetica', 8)
        self.setFillColor(colors.HexColor('#64748B'))
        ctr_info = getattr(self, 'contract_info', 'Contract-Cadru Prestări Servicii Contabilitate')
        self.drawString(36, 26, ctr_info)
        page_str = f"Pagina {self._pageNumber} din {page_count}"
        self.drawRightString(w - 36, 26, page_str)
        self.restoreState()

def generate_pdf_contract(prestator_key, client_data, nr_ctr, data_ctr, pret_baza=500, trans_incluse="0 - 40", data_start=None):
    prest = PRESTATORI_DATA.get(prestator_key, PRESTATORI_DATA["RIGHT WAY - TAX CONSULTING & MORE SRL"])
    data_start = data_start or data_ctr

    styles = getSampleStyleSheet()
    
    # Carefully calibrated font sizes to fit 3 pages for main body + signatures
    p_title = ParagraphStyle('CTitle', fontName='Helvetica-Bold', fontSize=12, leading=15, alignment=1, textColor=colors.HexColor('#1B365D'))
    p_sub = ParagraphStyle('CSub', fontName='Helvetica', fontSize=8.5, leading=11.5, alignment=1, textColor=colors.HexColor('#475569'))
    p_h1 = ParagraphStyle('CH1', fontName='Helvetica-Bold', fontSize=9.5, leading=12.5, textColor=colors.HexColor('#1B365D'), spaceBefore=5, spaceAfter=2.5, keepWithNext=True)
    p_h2 = ParagraphStyle('CH2', fontName='Helvetica-Bold', fontSize=8.5, leading=11.5, textColor=colors.HexColor('#2B4C7E'), spaceBefore=4, spaceAfter=2, keepWithNext=True)
    p_body = ParagraphStyle('CBody', fontName='Helvetica', fontSize=7.8, leading=10.2, textColor=colors.HexColor('#1E293B'), spaceBefore=1, spaceAfter=1.8, alignment=4)
    p_crit = ParagraphStyle('CCrit', fontName='Helvetica-Bold', fontSize=7.8, leading=10.2, textColor=colors.HexColor('#991B1B'), spaceBefore=1.5, spaceAfter=2, alignment=4)
    p_table = ParagraphStyle('CTable', fontName='Helvetica', fontSize=7.2, leading=9.2, textColor=colors.HexColor('#1E293B'))
    p_table_h = ParagraphStyle('CTableH', fontName='Helvetica-Bold', fontSize=7.2, leading=9.2, textColor=colors.white, alignment=1)

    story = []

    # Title
    story.append(Spacer(1, 2))
    story.append(Paragraph("<b>CONTRACT-CADRU DE PRESTĂRI SERVICII<br/>FINANCIAR-CONTABILE ȘI RESURSE UMANE</b>", p_title))
    story.append(Spacer(1, 1))
    story.append(Paragraph(f"<i>Nr. de înregistrare: <b>{nr_ctr}</b> &nbsp;|&nbsp; Data încheierii: <b>{data_ctr}</b></i>", p_sub))
    story.append(Spacer(1, 2))

    # Preambul
    story.append(Paragraph(
        "Încheiat în temeiul art. 14 din Ordonanța Guvernului nr. 65/1994 privind organizarea activității de expertiză contabilă "
        "și a contabililor autorizați, republicată, a Noului Cod Civil român privind contractul de mandat cu reprezentare și "
        "prestările de servicii, precum și a Regulamentului (UE) 2016/679 (RGPD), între următoarele părți contractante:",
        p_body
    ))
    story.append(Spacer(1, 2))

    # Cap I - Părți
    story.append(Paragraph("CAPITOLUL I. PĂRȚILE CONTRACTANTE", p_h1))
    col_w = 260
    t_prest = f"""<b>1. PRESTATORUL:</b><br/>
<b>{prest['nume']}</b><br/>
Sediu: {prest['sediu']}<br/>
Reg. Com.: {prest['reg_com']} &nbsp;|&nbsp; CUI: <b>{prest['cui']}</b><br/>
{prest['ceccar']}<br/>
Cont IBAN: <b>{prest['iban']}</b><br/>
Banca: {prest['banca']}<br/>
Reprezentant: <b>{prest['reprezentant']}</b> ({prest['functie']})<br/>
E-mail: {prest['email']} &nbsp;|&nbsp; Tel: {prest['telefon']}"""

    t_benef = f"""<b>2. BENEFICIARUL (CLIENTUL):</b><br/>
<b>{client_data.get('nume', '____________________ S.R.L.')}</b><br/>
Sediu: {client_data.get('sediu', '____________________')}<br/>
Reg. Com.: {client_data.get('reg_com', 'J___/______/________')} &nbsp;|&nbsp; CUI: <b>{client_data.get('cui', '________')}</b><br/>
Cont IBAN: <b>{client_data.get('iban', 'RO____________________')}</b><br/>
Banca: {client_data.get('banca', '____________________')}<br/>
Reprezentant: <b>{client_data.get('reprezentant', '____________________')}</b> ({client_data.get('functie', 'Administrator')})<br/>
E-mail: {client_data.get('email', '____________________')} &nbsp;|&nbsp; Tel: {client_data.get('telefon', '__________')}"""
    tbl_parti = Table([[Paragraph(t_prest, p_table), Paragraph(t_benef, p_table)]], colWidths=[col_w, col_w])
    tbl_parti.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (0,0), colors.HexColor('#F1F5F9')),
        ('BACKGROUND', (1,0), (1,0), colors.HexColor('#F8FAFC')),
        ('BOX', (0,0), (-1,-1), 0.6, colors.HexColor('#CBD5E1')),
        ('INNERGRID', (0,0), (-1,-1), 0.6, colors.HexColor('#CBD5E1')),
        ('TOPPADDING', (0,0), (-1,-1), 4),
        ('BOTTOMPADDING', (0,0), (-1,-1), 4),
        ('LEFTPADDING', (0,0), (-1,-1), 5),
        ('RIGHTPADDING', (0,0), (-1,-1), 5),
    ]))
    story.append(tbl_parti)
    story.append(Spacer(1, 2))

    # Cap II - Obiect
    story.append(Paragraph("CAPITOLUL II. OBIECTUL CONTRACTULUI", p_h1))
    story.append(Paragraph("<b>Art. 1. – Obiectul general:</b> Prestatorul se obligă să asigure în favoarea Beneficiarului servicii profesionale de evidență financiar-contabilă completă, asistență și consultanță fiscală primară, precum și servicii de evidență a personalului și salarizare (resurse umane), detaliate exhaustiv în <b>Anexa nr. 1</b>, parte integrantă din prezentul Contract.", p_body))
    story.append(Paragraph("<b>Art. 2. – Standarde profesionale și subcontractare:</b> Serviciile ce fac obiectul Contractului vor fi prestate direct de către personalul de specialitate al Prestatorului și/sau subcontractate către terți specialiști autorizați CECCAR/CCF, Prestatorul menținându-și răspunderea profesională deplină în raport cu Beneficiarul. Toate lucrările se execută conform standardelor profesionale emise de CECCAR și legislației fiscale în vigoare.", p_body))
    story.append(Paragraph("<b>Art. 3. – Reprezentare și emitere instrucțiuni:</b> Beneficiarul confirmă că reprezentantul său legal (administratorul sau persoana împuternicită expres în scris) are competența deplină de a emite instrucțiuni oficiale de lucru către Prestator, de a aproba statele de salarii, raportările și înregistrările contabile pe canalele electronice agreate.", p_body))
    story.append(Paragraph("<b>Art. 4. – Structura Anexelor:</b> Prezentul Contract se completează de drept cu: Anexa nr. 1 (Pachetul de bază de servicii de contabilitate, fiscalitate și salarizare); Anexa nr. 2 (Grila de tarife lunare și lista tarifelor pentru activități suplimentare 2026); Anexa nr. 3 (Procedura privind accesul și transmiterea în REGES-ONLINE și protecția datelor).", p_body))

    # Cap III - Durata
    story.append(Paragraph("CAPITOLUL III. DURATA CONTRACTULUI", p_h1))
    story.append(Paragraph(f"<b>Art. 5. – Durata inițială:</b> Prezentul Contract se încheie pe o durată fermă de 12 (douăsprezece) luni, intrând în vigoare la data semnării sale (sau la data de {data_start}).", p_body))
    story.append(Paragraph("<b>Art. 6. – Prelungire tacită:</b> Durata contractului se prelungește automat și succesiv pe noi perioade egale de câte 12 luni, cu excepția cazului în care una dintre Părți notifică în scris celeilalte Părți intenția de denunțare sau neprelungire, cu cel puțin 30 (treizeci) de zile calendaristice înainte de expirarea perioadei în curs.", p_body))

    # Cap IV - Pret
    story.append(Paragraph("CAPITOLUL IV. PREȚUL CONTRACTULUI, TARIFE ȘI MODALITĂȚI DE PLATĂ", p_h1))
    story.append(Paragraph(f"<b>Art. 7. – Prețul de bază și indexarea conform grilei:</b> Pentru serviciile curente de contabilitate, Beneficiarul va achita Prestatorului un onorariu lunar stabilit pe baza volumului de activitate. <b>Onorariul minim de bază este de {pret_baza} lei/lună pentru tranșa de {trans_incluse} tranzacții pentru microîntreprinderi fără TVA conform grilei din Anexa 2.</b> În situația în care volumul lunar de documente și tranzacții prelucrate depășește tranșa inițială, onorariul lunar se va ajusta automat corespunzător volumului efectiv prelucrat și regimului fiscal (neplătitor/plătitor de TVA, declarații speciale D390/profit), conform Grilei Oficiale de Prețuri din Anexa nr. 2. În cazul firmelor fără activitate tariful este de 400 lei.", p_body))
    story.append(Paragraph("<b>Art. 8. – Raportări financiare (Bilanțuri anuale și interimare):</b> Întocmirea, verificarea și depunerea situațiilor financiare anuale (bilanț) și a raportărilor contabile semestriale/interimare nu sunt incluse în onorariul lunar curent. Acestea se tarifează distinct cu contravaloarea unei luni de contabilitate (dar nu mai puțin de 400 lei pentru situațiile interimare), achitabile cu cel puțin 15-30 de zile înainte de termenul legal de depunere. Întocmirea și depunerea raportărilor se realizează după achitarea efectivă a tarifului aferent.", p_body))
    story.append(Paragraph("<b>Art. 9. – Servicii suplimentare și salarizare:</b> Serviciile de salarizare și administrare personal (50 lei/salariat/lună), precum și orice activități speciale solicitate la cerere (Revisal, dosare TVA, bănci suplimentare, certificate ONRC, asistență controale etc.) se tarifează separat conform Anexei nr. 2 (Secțiunea III) și se evidențiază ca poziții distincte pe factura lunară.", p_body))
    story.append(Paragraph("<b>Art. 10. – Facturare, termene de plată și penalități:</b> Factura fiscală se emite lunar de către Prestator în primele 5 (cinci) zile ale lunii următoare prestării serviciilor. Plata devine scadentă în termen de 5 zile calendaristice de la emitere. În cazul depășirii scadenței cu mai mult de 5 zile lucrătoare de grație, Prestatorul poate percepe penalități de întârziere de 0,1% pe zi din suma restantă. Dacă restanțele depășesc 30 de zile, Prestatorul poate suspenda prestarea serviciilor cu exonerare de răspundere și poate rezilia contractul de plin drept. Prețurile sunt nete și nu conțin TVA.", p_body))

    # Cap V - Predare & Intarzieri
    story.append(Paragraph("CAPITOLUL V. PREDAREA DOCUMENTELOR, TERMENE ȘI REGIMUL ÎNTÂRZIERILOR", p_h1))
    story.append(Paragraph("<b>Art. 11. – Termenul standard optim de predare:</b> Beneficiarul are obligația de a preda sau de a încărca electronic în folderul securizat Drive / platforma dedicată de lucru totalitatea documentelor financiar-contabile justificative aferente lunii precedente în intervalul <b>01 - 10 ale lunii în curs</b>. Respectarea acestui interval asigură prelucrarea riguroasă și depunerea fără costuri suplimentare a declarațiilor fiscale până la scadența legală de 25 ale lunii.", p_body))
    story.append(Paragraph("<b>Art. 12. – Regimul predărilor întârziate (Majorări pentru prelucrare în regim de urgență):</b> În cazul în care Beneficiarul predă documentele cu depășirea termenului standard de 10 ale lunii, pentru compensarea efortului operativ și alocarea de resurse suplimentare de urgență, se pot aplica următoarele ajustări:<br/>• Predare în intervalul 11 - 18 ale lunii: onorariul lunar curent se majorează cu <b>15%</b> pentru prelucrarea în regim prioritar, garantându-se depunerea declarațiilor la scadența legală;<br/>• Predare în intervalul 19 - 23 ale lunii: onorariul lunar curent se majorează cu <b>30%</b> pentru prelucrare în regim de urgență maximă, depunerea declarațiilor realizându-se la limită.", p_body))
    story.append(Paragraph("<b>Art. 13. – CLAUZĂ CRITICĂ PRIVIND TERMENUL LIMITĂ DE 23 ALE LUNII (EXONERARE DE RĂSPUNDERE): În situația în care documentele justificative NU sunt predate sau încărcate complet de către Beneficiar până cel târziu la data de 23 ale lunii, PRESTATORUL NU ÎȘI ASUMĂ OBLIGAȚIA ȘI RESPONSABILITATEA DEPUNERII DECLARAȚIILOR FISCALE AFERENTE ÎN CURSUL LUNII RESPECTIVE (până la termenul legal de 25). În acest caz, documentele vor fi procesate în ordinea normală a fluxului de lucru, iar declarațiile vor fi depuse după termen sau rectificate conform legii. Prestatorul este de plin drept EXONERAT de orice răspundere juridică, fiscală sau materială, iar toate eventualele consecințe negative — incluzând amenzi, dobânzi, penalități de întârziere ANAF, decizii de impunere din oficiu sau popriri — cad în mod exclusiv și irevocabil în sarcina Beneficiarului. În cazul declarațiilor rectificate se va percepe cost suplimentar conform Anexei 2.</b>", p_crit))
    story.append(Paragraph("<b>Art. 14. – Rapoarte și predarea documentelor prelucrate:</b> Prestatorul pune la dispoziția Beneficiarului, pe platforma securizată de lucru, documentele rezultate în urma prelucrării: balanțele lunare de verificare, jurnalele de cumpărări și vânzări, statele de plată a salariilor, sumele care trebuiesc plătite pentru buget. După încheierea anului fiscal vor fi încărcate declarațiile și confirmările recipiselor de depunere a declarațiilor fiscale.", p_body))

    # Cap VI - Drepturi & Obligatii
    story.append(Paragraph("CAPITOLUL VI. DREPTURILE ȘI OBLIGAȚIILE PĂRȚILOR", p_h1))
    story.append(Paragraph("<b>Art. 15. – Obligațiile esențiale ale Prestatorului:</b> a) Să asigure înregistrarea cronologică și sistematică a documentelor în partidă dublă; b) Să întocmească jurnalele contabile, balanțele lunare de verificare și registrele obligatorii conform legii; c) Să calculeze obligațiile fiscale și să transmită declarațiile fiscale în termenele legale, sub condiția primirii documentelor complete în termenele contractuale; d) Să asigure evidența salariaților și operarea instrucțiunilor primite în REGES-ONLINE conform Anexei nr. 3; e) Să informeze Beneficiarul asupra modificărilor legislative relevante cu impact asupra activității sale; f) Să asigure confidențialitatea strictă a datelor și informațiilor primite.", p_body))
    story.append(Paragraph("<b>Art. 16. – Obligațiile esențiale ale Beneficiarului:</b> a) Să pună la dispoziția Prestatorului documente primare complete, corecte, lizibile și legale, exclusiv aferente activității economice a societății; b) Să respecte termenele de predare a documentelor (01 - 10 ale lunii); c) Să achite onorariul convenit și contravaloarea serviciilor suplimentare la termenele scadente; d) Să asigure accesul Prestatorului în SPV (Spațiul Privat Virtual) și să transmită extrasele bancare complete în format electronic; e) Să comunice în scris și în timp util orice modificare privind salariații (angajări, încetări, concedii medicale etc.); f) Să furnizeze orice lămuriri suplimentare solicitate de Prestator pentru reflectarea fidelă a tranzacțiilor.", p_body))

    # Cap VII - Delimitare Raspundere
    story.append(Paragraph("CAPITOLUL VII. DELIMITAREA RĂSPUNDERII ȘI LIMITAREA DESPĂGUBIRILOR", p_h1))
    story.append(Paragraph("<b>Art. 17. – Delimitarea răspunderii patrimoniale:</b> Prestatorul răspunde exclusiv de reflectarea corectă în contabilitate a operațiunilor economice și de calculul corect al taxelor, pe baza documentelor justificative puse la dispoziție de Beneficiar. Prestatorul se obligă să suporte amenzile și penalitățile aplicate de organele fiscale rezultate exclusiv din culpa sa profesională dovedită. Prestatorul NU răspunde în niciun fel de: a) Gestiunea fizică a patrimoniului, recepția bunurilor, lipsurile din inventar sau efectuarea inventarierilor anuale; b) Legalitatea, autenticitatea și oportunitatea economică a documentelor emise de partenerii comerciali ai Beneficiarului; c) Emiterea facturilor proprii și transmiterea lor în sistemul național RO e-Factura; d) Deducerea integrală (100%) a cheltuielilor și TVA auto în lipsa foilor de parcurs complete și corecte puse la dispoziție de Beneficiar; e) Calculul și plata taxelor și impozitelor locale (clădiri, autoturisme, terenuri); f) Modul în care a fost ținută contabilitatea anterior preluării mandatului de către Prestator sau de întârzierile cauzate de refacerea evidențelor anterioare.", p_body))
    story.append(Paragraph("<b>Art. 18. – Plafonarea despăgubirilor:</b> Răspunderea profesională totală a Prestatorului pentru eventuale prejudicii materiale directe cauzate Beneficiarului exclusiv din culpa sa profesională dovedită este expres și irevocabil plafonată la maximum contravaloarea a <b>3 (trei) onorarii lunare contractuale</b> de contabilitate.", p_body))
    story.append(Paragraph("<b>Art. 19. – Dreptul de retenție asupra bazei de date electronice:</b> Prestatorul își rezervă dreptul legitim de retenție asupra bazei de date contabile generate în softul de specialitate (ex: Saga) și asupra dosarelor contabile până la achitarea integrală de către Beneficiar a tuturor onorariilor și debitelor restante. În lipsa plății integrale, baza de date contabilă se conservă pentru maximum 6 luni de la încetarea contractului, după care poate fi ștearsă din sistemele Prestatorului fără altă notificare prealabilă.", p_body))
    story.append(Paragraph("<b>Art. 20. – Clauza de non-racolare a personalului:</b> Beneficiarul se obligă ferm să nu recruteze, să nu angajeze și să nu contracteze direct sau indirect servicii cu personalul angajat sau colaboratorii de specialitate ai Prestatorului, pe durata derulării prezentului contract și pe o perioadă de 12 luni de la încetarea acestuia. În cazul încălcării acestei obligații, Beneficiarul va achita Prestatorului o despăgubire forfetară fermă egală cu <b>contravaloarea a 24 de luni de onorariu</b> contractual de contabilitate.", p_body))

    # Cap VIII - RGPD
    story.append(Paragraph("CAPITOLUL VIII. PROTECȚIA DATELOR CU CARACTER PERSONAL (RGPD - REGULAMENTUL UE 2016/679)", p_h1))
    story.append(Paragraph("<b>Art. 21. – Cadrul general și conformitatea:</b> Părțile se obligă să respecte strict normele și cerințele stabilite de Regulamentul (UE) 2016/679 (RGPD), precum și legislația națională conexă (Legea nr. 190/2018). Prevederile prezentului capitol se aplică tuturor operațiunilor de colectare, înregistrare, organizare, stocare, consultare, transmitere sau arhivare a datelor cu caracter personal efectuate în executarea contractului.", p_body))
    story.append(Paragraph("<b>Art. 22. – Delimitarea rolurilor și calității părților (Art. 28 RGPD):</b> Beneficiarul are calitatea de <b>Operator de date cu caracter personal</b>; Prestatorul are calitatea de <b>Persoană Împuternicită de operator (Împuternicit / Procesator)</b>, prelucrând datele cu caracter personal exclusiv în numele, pe seama și în baza instrucțiunilor scrise ale Beneficiarului.", p_body))
    story.append(Paragraph("<b>Art. 23. – Categoriile de date și persoanele vizate:</b> Prelucrarea vizează datele cu caracter personal ale: salariaților, foștilor salariați și candidaților Beneficiarului, administratorilor, asociaților, precum și ale partenerilor comerciali persoane fizice (nume, prenume, CNP, serie/număr act identitate, adresă, cont bancar, funcție, pontaje, drepturi salariale, concedii medicale etc.).", p_body))
    story.append(Paragraph("<b>Art. 24. – Obligațiile Beneficiarului (Operatorul):</b> Garantează că deține temeiul juridic necesar pentru prelucrare și transmitere; asigură informarea persoanelor vizate; transmite instrucțiuni clare scrise conforme legii.", p_body))
    story.append(Paragraph("<b>Art. 25. – Obligațiile Prestatorului (Persoana Împuternicită):</b> Prelucrează datele numai pe baza instrucțiunilor Operatorului; asigură confidențialitatea personalului autorizat; implementează măsuri tehnice și organizatorice adecvate de securitate; notifică incidentele de securitate în maximum 48 de ore; pune la dispoziție informațiile pentru demonstrarea conformității.", p_body))
    story.append(Paragraph("<b>Art. 26. – Sub-împuterniciți și infrastructură IT:</b> Operatorul acordă Prestatorului autorizație generală de a utiliza furnizori IT de încredere (Saga Software, servicii cloud securizate Google Workspace/Google Drive, certificate de semnătură electronică calificată).", p_body))
    story.append(Paragraph("<b>Art. 27. – Regimul datelor la încetarea contractului:</b> După stingerea tuturor obligațiilor financiare, Prestatorul va preda Beneficiarului baza de date și documentele justificative. Arhivarea legală se supune termenelor obligatorii din Legea contabilității nr. 82/1991.", p_body))

    # Cap IX - Confidentialitate
    story.append(Paragraph("CAPITOLUL IX. CONFIDENȚIALITATEA COMERCIALĂ", p_h1))
    story.append(Paragraph("<b>Art. 28. – Clauza de confidențialitate comercială:</b> Toate informațiile contabile, financiare, comerciale și deciziile manageriale sunt strict confidențiale. Clauza rămâne în vigoare pe toată durata contractului și pentru o perioadă de <b>2 (doi) ani după încetarea acestuia</b>.", p_body))

    # Cap X - Forta Majora, Incetare, Litigii
    story.append(Paragraph("CAPITOLUL X. FORȚA MAJORĂ, ÎNCETAREA CONTRACTULUI ȘI LITIGII", p_h1))
    story.append(Paragraph("<b>Art. 29. – Forța majoră:</b> Forța majoră apără de răspundere partea care o invocă, cu notificare în 5 zile. Dacă se prelungește peste 30 de zile, contractul încetează de drept fără daune-interese.", p_body))
    story.append(Paragraph("<b>Art. 30. – Modalități de încetare a contractului:</b> a) Prin acordul scris al ambelor părți; b) Prin denunțare unilaterală cu preaviz scris de 30 de zile; c) De plin drept, prin reziliere la neplată de peste 30 de zile; d) De plin drept, la insolvență/faliment.", p_body))
    story.append(Paragraph("<b>Art. 31. – Soluționarea litigiilor:</b> Neînțelegerile se soluționează amiabil sau prin conciliere la Comisia de Arbitraj a filialei CECCAR competente. În caz contrar, litigiile revin instanțelor de la sediul Prestatorului.", p_body))

    # Cap XI - Comunicare
    story.append(Paragraph("CAPITOLUL XI. DATE DE CONTACT ȘI COMUNICARE OPERATIVĂ", p_h1))
    story.append(Paragraph(f"<b>Art. 32. – Comunicări oficiale:</b> Toate notificările se transmit prin e-mail la: Prestator: {prest['email']} (Tel: {prest['telefon']}); Beneficiar: {client_data.get('email', '____________________')} (Tel: {client_data.get('telefon', '__________')}).", p_body))

    # Semnaturi
    story.append(Spacer(1, 2))
    story.append(Paragraph(f"Încheiat și semnat astăzi, <b>{data_ctr}</b>, în 2 (două) exemplare originale cu valoare juridică egală.", p_body))
    story.append(Spacer(1, 2))

    txt_semn_prest = f"""<b>PRESTATOR:</b><br/>
<b>{prest['nume']}</b><br/>
Prin Administrator,<br/>
<b>{prest['reprezentant']}</b><br/><br/>
Semnătura: ___________________________<br/>
Ștampila:"""

    txt_semn_benef = f"""<b>BENEFICIAR:</b><br/>
<b>{client_data.get('nume', '____________________ S.R.L.')}</b><br/>
Prin Administrator,<br/>
<b>{client_data.get('reprezentant', '____________________')}</b><br/><br/>
Semnătura: ___________________________<br/>
Ștampila:"""

    t_semn = Table([[Paragraph(txt_semn_prest, p_table), Paragraph(txt_semn_benef, p_table)]], colWidths=[col_w, col_w])
    t_semn.setStyle(TableStyle([
        ('BOX', (0,0), (-1,-1), 0.6, colors.HexColor('#94A3B8')),
        ('INNERGRID', (0,0), (-1,-1), 0.6, colors.HexColor('#CBD5E1')),
        ('TOPPADDING', (0,0), (-1,-1), 4),
        ('BOTTOMPADDING', (0,0), (-1,-1), 4),
        ('LEFTPADDING', (0,0), (-1,-1), 6),
        ('RIGHTPADDING', (0,0), (-1,-1), 6),
    ]))
    story.append(t_semn)

    # ANEXA 1
    story.append(PageBreak())
    story.append(Paragraph("ANEXA NR. 1: PRESTAȚII DE BAZĂ (PACHET SERVICII CONTABILITATE ȘI SALARIZARE)", p_h1))
    story.append(Paragraph(f"<i>La Contractul-Cadru de Prestări Servicii nr. {nr_ctr} din data de {data_ctr}</i>", p_sub))
    story.append(Spacer(1, 3))
    story.append(Paragraph("<b>A. Evidență Financiar-Contabilă și Registre Obligatorii:</b>", p_h2))
    story.append(Paragraph("1. Contabilizarea cronologică și sistematică a documentelor justificative în partidă dublă (intrări, ieșiri, bancă, casă, deconturi).<br/>2. Evidența analitică și sintetică a clienților, furnizorilor, trezoreriei și imobilizărilor.<br/>3. Calculul amortizării contabile și fiscale lunare a mijloacelor fixe.<br/>4. Întocmirea balanțelor lunare de verificare analitice și sintetice (patru serii de egalități).<br/>5. Întocmirea registrelor contabile obligatorii: Registrul Jurnal, Registrul Inventar, Jurnal de Cumpărări, Jurnal de Vânzări, Registru de Evidență Fiscală.<br/>6. Punerea la dispoziție în format electronic a fișelor de cont și a Cărții Mari (la cerere).", p_body))

    story.append(Paragraph("<b>B. Salarizare și Administrare Personal (Pachet de bază):</b>", p_h2))
    story.append(Paragraph("1. Întocmirea lunară a statelor de plată pe baza pontajelor și certificatelor medicale.<br/>2. Calculul reținerilor salariale, CAS, CASS, CAM și a impozitului pe salarii.<br/>3. Calculul concediilor medicale și al impozitului pe tichete de masă/cadou/vacanță.<br/>4. Întocmirea ordinelor de plată pentru taxele salariale la bugetul de stat.<br/>5. Eliberarea adeverințelor tip pentru medic sau spitale la solicitarea salariaților.", p_body))

    story.append(Paragraph("<b>C. Declarații Fiscale și Raportări Periodice către Autorități:</b>", p_h2))
    story.append(Paragraph("1. Calculul impozitului pe veniturile microîntreprinderilor / profit, dividende și rețineri la sursă.<br/>2. Decontul de TVA (Formularul 300 / 301).<br/>3. Declarația informativă privind livrările/achizițiile naționale (Formularul 394).<br/>4. Declarația recapitulativă privind livrările/achizițiile intracomunitare (Formularul 390 VIES).<br/>5. Declarația privind obligațiile de plată la bugetul de stat (Formularul 100).<br/>6. Declarația privind contribuțiile sociale și evidența asiguraților (Formularul 112).<br/>7. Fișierul Standard de Control Fiscal (SAF-T / Declarația 406), conform legii (fără stocuri/producție).<br/>8. Declarația de mențiuni vector fiscal (Formularul 700) pe baza documentelor primite cu cel puțin 10 zile înainte.", p_body))

    # ANEXA 2
    story.append(PageBreak())
    story.append(Paragraph("ANEXA NR. 2: GRILA OFICIALĂ DE PREȚURI ȘI TARIFE SUPLIMENTARE (2026)", p_h1))
    story.append(Paragraph(f"<i>La Contractul-Cadru de Prestări Servicii nr. {nr_ctr} din data de {data_ctr}</i>", p_sub))
    story.append(Spacer(1, 3))

    story.append(Paragraph("<b>I. PENTRU ACTIVITĂȚI DE PRESTĂRI SERVICII</b>", p_h2))
    hdr_s = ["Nr. tranzacții lunare", "0 - 40", "40 - 100", "101 - 200", "201 - 300", "301 - 400", "401 - 500"]
    rows_s = [
        ["Microîntreprindere, fără TVA", "500 lei", "600 lei", "700 lei", "800 lei", "900 lei", "1.000 lei"],
        ["Cu TVA (D300 și D394)", "600 lei", "700 lei", "800 lei", "900 lei", "1.000 lei", "1.100 lei"],
        ["Cu D390 (servicii intrastat)", "650 lei", "750 lei", "850 lei", "950 lei", "1.050 lei", "1.150 lei"],
        ["Impozit pe profit", "700 lei", "800 lei", "900 lei", "1.000 lei", "1.100 lei", "1.200 lei"],
        ["Consultanță inclusă", "15 min", "15 min", "15 min", "15 min", "30 min", "30 min"]
    ]
    t_data_s = [[Paragraph(f"<b>{c}</b>", p_table_h) for c in hdr_s]]
    for r in rows_s:
        t_data_s.append([Paragraph(r[0], p_table)] + [Paragraph(c, p_table) for c in r[1:]])
    
    cw = [135] + [64]*6
    tbl_s = Table(t_data_s, colWidths=cw)
    tbl_s.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#1B365D')),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#CBD5E1')),
        ('TOPPADDING', (0,0), (-1,-1), 2.5),
        ('BOTTOMPADDING', (0,0), (-1,-1), 2.5),
    ]))
    story.append(tbl_s)
    story.append(Paragraph("<i>* Peste 500 de tranzacții lunare - prețul se negociază direct prin act adițional.</i>", p_sub))
    story.append(Spacer(1, 3))

    story.append(Paragraph("<b>II. PENTRU ACTIVITĂȚI DE MARFĂ / GESTIUNE</b>", p_h2))
    hdr_m = ["Nr. tranzacții lunare", "6 - 20", "21 - 100", "101 - 200", "201 - 300", "301 - 400", "401 - 500"]
    rows_m = [
        ["Microîntreprindere, fără TVA", "700 lei", "800 lei", "900 lei", "1.000 lei", "1.100 lei", "1.200 lei"],
        ["Cu TVA (D300 și D394)", "1.000 lei", "1.100 lei", "1.200 lei", "1.300 lei", "1.400 lei", "1.500 lei"],
        ["Cu D390 (servicii intrastat)", "1.050 lei", "1.150 lei", "1.250 lei", "1.350 lei", "1.450 lei", "1.550 lei"],
        ["Impozit pe profit", "1.150 lei", "1.250 lei", "1.350 lei", "1.450 lei", "1.550 lei", "1.650 lei"],
        ["Consultanță inclusă", "15 min", "15 min", "15 min", "15 min", "30 min", "30 min"]
    ]
    t_data_m = [[Paragraph(f"<b>{c}</b>", p_table_h) for c in hdr_m]]
    for r in rows_m:
        t_data_m.append([Paragraph(r[0], p_table)] + [Paragraph(c, p_table) for c in r[1:]])
    tbl_m = Table(t_data_m, colWidths=cw)
    tbl_m.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#2B4C7E')),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#CBD5E1')),
        ('TOPPADDING', (0,0), (-1,-1), 2.5),
        ('BOTTOMPADDING', (0,0), (-1,-1), 2.5),
    ]))
    story.append(tbl_m)
    story.append(Paragraph("<i>* Peste 500 de tranzacții lunare - prețul se negociază direct prin act adițional.</i>", p_sub))
    story.append(Spacer(1, 3))

    # Anexa 2 - Secțiunea III: Tarife suplimentare
    story.append(Paragraph("<b>III. TARIFE SUPLIMENTARE (ACTIVITĂȚI LA CERERE 2026)</b>", p_h2))
    suplim = [
        ("De la a 3-a bancă / cont bancar", "50 lei/bancă"),
        ("De la al 2-lea contract de leasing", "10 lei/lună"),
        ("Deschidere Revisal - primul contract", "200 lei"),
        ("Calcul salariu, stat de plată, pontaj, D112, OP-uri", "50 lei/salariat/lună"),
        ("Statistica", "începând de la 350 lei"),
        ("Obținere cod TVA la înființare / intracomunitar", "300 lei"),
        ("Obținere număr EORI (vamă)", "300 lei"),
        ("Dosar recuperare indemnizații FNUASS", "200 lei"),
        ("Obținere NIF / Certificat rezidență fiscală", "200 lei / 300 lei"),
        ("Depunere declarație casă de marcat", "50 lei"),
        ("Dosar depășire plafon de scutire TVA", "300 lei"),
        ("Înregistrare punct de lucru ANAF (fără / cu salariați)", "100 lei"),
        ("Întocmire Regulament Intern", "începând de la 350 lei"),
        ("Achiziționare Registru Unic de Control (RUC)", "începând de la 100 lei (+ cost registru)"),
        ("Asistență la rambursare TVA", "începând de la 500 lei"),
        ("Reactivare firmă / Activare cod TVA anulat", "300 lei"),
        ("Completare acte solicitate de bănci", "începând de la 300 lei"),
        ("Certificat constatator ONRC", "130 lei"),
        ("Consultanță fiscală suplimentară", "50 EURO/oră"),
        ("Secretariat - aranjare documente (până la 50 doc.)", "50 lei"),
        ("Bilanț interimar / Situații financiare anuale", "costul unei luni (minim 350 lei)"),
        ("Întocmire manual de politici și proceduri contabile", "500 lei")
    ]
    t_data_sup = [[Paragraph("<b>Serviciu / Activitate specifică</b>", p_table_h), Paragraph("<b>Tarif 2026</b>", p_table_h)]]
    for s, t in suplim:
        t_data_sup.append([Paragraph(s, p_table), Paragraph(t, p_table)])
    tbl_sup = Table(t_data_sup, colWidths=[360, 159])
    tbl_sup.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#1B365D')),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#CBD5E1')),
        ('TOPPADDING', (0,0), (-1,-1), 2),
        ('BOTTOMPADDING', (0,0), (-1,-1), 2),
    ]))
    story.append(tbl_sup)
    story.append(Paragraph("<i>Prețurile nu includ taxe/costuri oficiale plătite către instituții (taxe ANAF, notariale, poștale, ONRC).</i>", p_sub))

    # ANEXA 3
    story.append(PageBreak())
    story.append(Paragraph("ANEXA NR. 3: PROCEDURA PRIVIND ACCESUL ȘI TRANSMITEREA ÎN REGES-ONLINE ȘI PROTECȚIA DATELOR", p_h1))
    story.append(Paragraph(f"<i>La Contractul-Cadru de Prestări Servicii nr. {nr_ctr} din data de {data_ctr}</i>", p_sub))
    story.append(Spacer(1, 3))
    story.append(Paragraph("<b>Art. 1. – Cadrul legal și crearea contului de acces:</b> În temeiul prevederilor stabilite prin „Procedura de acces în vederea completării, transmiterii și interogării datelor în/din Registrul general de evidență a salariaților - REGES-ONLINE” din 01.08.2025 a Ministerului Muncii, Familiei, Tineretului și Solidarității Sociale, Prestatorul, prin reprezentantul său legal – administrator, se obligă să întreprindă toate demersurile necesare pentru crearea și gestionarea contului de acces al angajatorului Beneficiar.", p_body))
    story.append(Paragraph("<b>Art. 2. – Persoana desemnată și respectarea RGPD:</b> Reprezentantul desemnat al Prestatorului este persoana împuternicită prin decizie scrisă, cu respectarea strictă a dispozițiilor Regulamentului (UE) 2016/679 (RGPD), pentru completarea, actualizarea și transmiterea datelor în Registru.", p_body))
    story.append(Paragraph("<b>Art. 3. – Instrucțiuni scrise de lucru:</b> Beneficiarul va transmite către Prestator instrucțiuni clare scrise prin e-mail (la adresa oficială declarată în contract) pentru orice operațiune legată de contractele de muncă: încheiere contract individual de muncă nou, modificare funcție/salariu/normă, suspendare, detașare sau încetare a raporturilor de muncă.", p_body))
    story.append(Paragraph("<b>Art. 4. – Termenele legale și răspunderea transmiterii:</b> Obligația transmiterii datelor și documentelor către Prestator anterior termenelor legale revine exclusiv Beneficiarului: pentru angajare – cel târziu în ziua anterioară începerii activității de către salariat; pentru orice alte modificări/încetări – cu respectarea termenelor prevăzute de Codul Muncii și normele metodologice. Prestatorul răspunde de înregistrarea și transmiterea corectă în sistemul REGES-ONLINE a datelor primite în formă scrisă în timp util.", p_body))
    story.append(Paragraph("<b>Art. 5. – Înregistrarea modificărilor salariale de către Beneficiar:</b> În cazul în care administrarea și calculul salariilor sunt asigurate în mod excepțional direct de către Beneficiar, acesta are obligația de a înștiința neîntârziat Prestatorul în scris pentru a opera corect notele contabile de salarii și obligațiile bugetare conexe.", p_body))
    story.append(Spacer(1, 6))

    # Semnaturi anexa 3
    t_semn_anexe = Table([[Paragraph(txt_semn_prest, p_table), Paragraph(txt_semn_benef, p_table)]], colWidths=[col_w, col_w])
    t_semn_anexe.setStyle(TableStyle([
        ('BOX', (0,0), (-1,-1), 0.6, colors.HexColor('#94A3B8')),
        ('INNERGRID', (0,0), (-1,-1), 0.6, colors.HexColor('#CBD5E1')),
        ('TOPPADDING', (0,0), (-1,-1), 4),
        ('BOTTOMPADDING', (0,0), (-1,-1), 4),
        ('LEFTPADDING', (0,0), (-1,-1), 6),
        ('RIGHTPADDING', (0,0), (-1,-1), 6),
    ]))
    story.append(t_semn_anexe)

    # Build Document
    pdf_buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        pdf_buffer,
        pagesize=A4,
        leftMargin=36,
        rightMargin=36,
        topMargin=60,
        bottomMargin=46
    )

    def make_canvas(*args, **kwargs):
        c = NumberedCanvasWithLogos(*args, **kwargs)
        c.prestator_logo = prest['logo_path']
        c.contract_info = f"Contract nr. {nr_ctr} / {data_ctr} | Beneficiar: {client_data.get('nume', 'CLIENT')}"
        return c

    doc.build(story, canvasmaker=make_canvas)
    return pdf_buffer.getvalue()

def set_cell_background(cell, hex_color):
    shading_xml = f'<w:shd {nsdecls("w")} w:fill="{hex_color}" w:val="clear"/>'
    cell._tc.get_or_add_tcPr().append(parse_xml(shading_xml))

def set_cell_margins(cell, top=60, bottom=60, left=80, right=80):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('w:top', top), ('w:bottom', bottom), ('w:left', left), ('w:right', right)]:
        node = OxmlElement(m)
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def add_field(p, text):
    run = p.add_run()
    fld = parse_xml(r'<w:fldSimple %s w:instr="%s"/>' % (nsdecls('w'), text))
    run._r.append(fld)

def generate_docx_contract(prestator_key, client_data, nr_ctr, data_ctr, pret_baza=500, trans_incluse="0 - 40", data_start=None):
    prest = PRESTATORI_DATA.get(prestator_key, PRESTATORI_DATA["RIGHT WAY - TAX CONSULTING & MORE SRL"])
    data_start = data_start or data_ctr

    doc = Document()
    # Set default style font to Helvetica
    style_normal = doc.styles['Normal']
    style_normal.font.name = 'Helvetica'
    style_normal.font.size = Pt(9.5)
    style_normal.font.color.rgb = RGBColor(30, 41, 59)
    for h_name in ['Heading 1', 'Heading 2', 'Heading 3']:
        if h_name in doc.styles:
            doc.styles[h_name].font.name = 'Helvetica' 
    for s in doc.sections:
        s.top_margin = Inches(0.7)
        s.bottom_margin = Inches(0.7)
        s.left_margin = Inches(0.7)
        s.right_margin = Inches(0.7)
        s.header.is_linked_to_previous = False
        s.footer.is_linked_to_previous = False

        # Header with logo images
        hp = s.header.paragraphs[0]
        hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        hr = hp.add_run(f"CONTRACT SERVICII CONTABILITATE | {prest['nume']}")
        hr.font.size = Pt(8.5)
        hr.font.color.rgb = RGBColor(100, 116, 139)

        # Footer
        fp = s.footer.paragraphs[0]
        fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        fr = fp.add_run(f"Contract nr. {nr_ctr} / {data_ctr} | Pagina ")
        fr.font.size = Pt(8.5)
        fr.font.color.rgb = RGBColor(100, 116, 139)
        add_field(fp, "PAGE")
        fr2 = fp.add_run(" din ")
        fr2.font.size = Pt(8.5)
        fr2.font.color.rgb = RGBColor(100, 116, 139)
        add_field(fp, "NUMPAGES")

    def set_p(p, b=1.5, a=3):
        p.paragraph_format.space_before = Pt(b)
        p.paragraph_format.space_after = Pt(a)
        p.paragraph_format.line_spacing = 1.15

    # Title
    p_t = doc.add_paragraph()
    set_p(p_t, 4, 1)
    p_t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_t = p_t.add_run("CONTRACT-CADRU DE PRESTĂRI SERVICII\nFINANCIAR-CONTABILE ȘI RESURSE UMANE")
    r_t.font.bold = True
    r_t.font.size = Pt(14)
    r_t.font.color.rgb = RGBColor(27, 54, 93)

    p_s = doc.add_paragraph()
    set_p(p_s, 0, 6)
    p_s.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_s = p_s.add_run(f"Nr. de înregistrare: {nr_ctr} / ___________ | Data încheierii: {data_ctr}")
    r_s.font.italic = True
    r_s.font.size = Pt(9.5)

    p_pre = doc.add_paragraph()
    set_p(p_pre, 2, 5)
    p_pre.add_run(
        "Încheiat în temeiul art. 14 din Ordonanța Guvernului nr. 65/1994 privind organizarea activității de expertiză contabilă "
        "și a contabililor autorizați, republicată, a Noului Cod Civil român privind contractul de mandat cu reprezentare și "
        "prestările de servicii, precum și a Regulamentului (UE) 2016/679 (RGPD), între următoarele părți contractante:"
    )

    # Cap I
    doc.add_heading("CAPITOLUL I. PĂRȚILE CONTRACTANTE", level=1)
    tbl = doc.add_table(rows=1, cols=2)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    tbl.rows[0].cells[0].width = Inches(3.5)
    tbl.rows[0].cells[0].width = Inches(3.5)

    cp = tbl.rows[0].cells[0]
    set_cell_background(cp, "F1F5F9")
    set_cell_margins(cp)
    pp = cp.paragraphs[0]
    rp = pp.add_run("1. PRESTATORUL\n")
    rp.font.bold = True
    rp.font.color.rgb = RGBColor(27, 54, 93)
    pp.add_run(
        f"{prest['nume']}\nSediu: {prest['sediu']}\nReg. Com.: {prest['reg_com']}\n"
        f"CUI: {prest['cui']}\n{prest['ceccar']}\nIBAN: {prest['iban']}\nBanca: {prest['banca']}\n"
        f"Reprezentant: {prest['reprezentant']} ({prest['functie']})\n"
        f"E-mail: {prest['email']} | Tel: {prest['telefon']}"
    )

    cb = tbl.rows[0].cells[1]
    set_cell_background(cb, "F8FAFC")
    set_cell_margins(cb)
    pb = cb.paragraphs[0]
    rb = pb.add_run("2. BENEFICIARUL (CLIENTUL)\n")
    rb.font.bold = True
    rb.font.color.rgb = RGBColor(27, 54, 93)
    pb.add_run(
        f"{client_data.get('nume', '____________________ S.R.L.')}\n"
        f"Sediu: {client_data.get('sediu', '____________________')}\n"
        f"Reg. Com.: {client_data.get('reg_com', 'J___/______/________')}\n"
        f"CUI: {client_data.get('cui', '________')}\n"
        f"IBAN: {client_data.get('iban', 'RO____________________')}\n"
        f"Banca: {client_data.get('banca', '____________________')}\n"
        f"Reprezentant: {client_data.get('reprezentant', '____________________')} ({client_data.get('functie', 'Administrator')})\n"
        f"E-mail: {client_data.get('email', '____________________')} | Tel: {client_data.get('telefon', '__________')}"
    )

    def add_a(num, titlu, text, bold_red=False):
        p = doc.add_paragraph()
        set_p(p, 1.5, 2.5)
        r = p.add_run(f"{num} ")
        r.font.bold = True
        if bold_red:
            r.font.color.rgb = RGBColor(153, 27, 27)
        else:
            r.font.color.rgb = RGBColor(27, 54, 93)
        if titlu:
            r2 = p.add_run(f"– {titlu}: ")
            r2.font.bold = True
            if bold_red:
                r2.font.color.rgb = RGBColor(153, 27, 27)
        r_txt = p.add_run(text)
        if bold_red:
            r_txt.font.bold = True
            r_txt.font.color.rgb = RGBColor(153, 27, 27)

    doc.add_heading("CAPITOLUL II. OBIECTUL CONTRACTULUI", level=1)
    add_a("Art. 1.", "Obiectul general", "Prestatorul se obligă să asigure în favoarea Beneficiarului servicii profesionale de evidență financiar-contabilă completă, asistență și consultanță fiscală primară, precum și servicii de evidență a personalului și salarizare (resurse umane), detaliate exhaustiv în Anexa nr. 1, parte integrantă din prezentul Contract.")
    add_a("Art. 2.", "Standarde profesionale și subcontractare", "Serviciile ce fac obiectul Contractului vor fi prestate direct de către personalul de specialitate al Prestatorului și/sau subcontractate către terți specialiști autorizați CECCAR/CCF, Prestatorul menținându-și răspunderea profesională deplină în raport cu Beneficiarul. Toate lucrările se execută conform standardelor profesionale emise de CECCAR și legislației fiscale în vigoare.")
    add_a("Art. 3.", "Reprezentare și emitere instrucțiuni", "Beneficiarul confirmă că reprezentantul său legal (administratorul sau persoana împuternicită expres în scris) are competența deplină de a emite instrucțiuni oficiale de lucru către Prestator, de a aproba statele de salarii, raportările și înregistrările contabile pe canalele electronice agreate.")
    add_a("Art. 4.", "Structura Anexelor", "Prezentul Contract se completează de drept cu: Anexa nr. 1 (Pachetul de bază de servicii de contabilitate, fiscalitate și salarizare); Anexa nr. 2 (Grila de tarife lunare și lista tarifelor pentru activități suplimentare 2026); Anexa nr. 3 (Procedura privind accesul și transmiterea în REGES-ONLINE și protecția datelor).")

    doc.add_heading("CAPITOLUL III. DURATA CONTRACTULUI", level=1)
    add_a("Art. 5.", "Durata inițială", f"Prezentul Contract se încheie pe o durată fermă de 12 (douăsprezece) luni, intrând în vigoare la data semnării sale (sau la data de {data_start}).")
    add_a("Art. 6.", "Prelungire tacită", "Durata contractului se prelungește automat și succesiv pe noi perioade egale de câte 12 luni, cu excepția cazului în care una dintre Părți notifică în scris celeilalte Părți intenția de denunțare sau neprelungire, cu cel puțin 30 (treizeci) de zile calendaristice înainte de expirarea perioadei în curs.")

    doc.add_heading("CAPITOLUL IV. PREȚUL CONTRACTULUI, TARIFE ȘI MODALITĂȚI DE PLATĂ", level=1)
    add_a("Art. 7.", "Prețul de bază și indexarea conform grilei", f"Pentru serviciile curente de contabilitate, Beneficiarul va achita Prestatorului un onorariu lunar stabilit pe baza volumului de activitate. Onorariul minim de bază este de {pret_baza} lei/lună pentru tranșa de {trans_incluse} tranzacții pentru microîntreprinderi fără TVA conform grilei din Anexa 2. În situația în care volumul lunar de documente și tranzacții prelucrate depășește tranșa inițială, onorariul lunar se va ajusta automat corespunzător volumului efectiv prelucrat și regimului fiscal (neplătitor/plătitor de TVA, declarații speciale D390/profit), conform Grilei Oficiale de Prețuri din Anexa nr. 2. În cazul firmelor fără activitate tariful este de 400 lei.")
    add_a("Art. 8.", "Raportări financiare (Bilanțuri anuale și interimare)", "Întocmirea, verificarea și depunerea situațiilor financiare anuale (bilanț) și a raportărilor contabile semestriale/interimare nu sunt incluse în onorariul lunar curent. Acestea se tarifează distinct cu contravaloarea unei luni de contabilitate (dar nu mai puțin de 400 lei pentru situațiile interimare), achitabile cu cel puțin 15-30 de zile înainte de termenul legal de depunere. Întocmirea și depunerea raportărilor se realizează după achitarea efectivă a tarifului aferent.")
    add_a("Art. 9.", "Servicii suplimentare și salarizare", "Serviciile de salarizare și administrare personal (50 lei/salariat/lună), precum și orice activități speciale solicitate la cerere (Revisal, dosare TVA, bănci suplimentare, certificate ONRC, asistență controale etc.) se tarifează separat conform Anexei nr. 2 (Secțiunea III) și se evidențiază ca poziții distincte pe factura lunară.")
    add_a("Art. 10.", "Facturare, termene de plată și penalități", "Factura fiscală se emite lunar de către Prestator în primele 5 (cinci) zile ale lunii următoare prestării serviciilor. Plata devine scadentă în termen de 5 zile calendaristice de la emitere. În cazul depășirii scadenței cu mai mult de 5 zile lucrătoare de grație, Prestatorul poate percepe penalități de întârziere de 0,1% pe zi din suma restantă. Dacă restanțele depășesc 30 de zile, Prestatorul poate suspenda prestarea serviciilor cu exonerare de răspundere și poate rezilia contractul de plin drept. Prețurile sunt nete și nu conțin TVA.")

    doc.add_heading("CAPITOLUL V. PREDAREA DOCUMENTELOR, TERMENE ȘI REGIMUL ÎNTÂRZIERILOR", level=1)
    add_a("Art. 11.", "Termenul standard optim de predare", "Beneficiarul are obligația de a preda sau de a încărca electronic în folderul securizat Drive / platforma dedicată de lucru totalitatea documentelor financiar-contabile justificative aferente lunii precedente în intervalul 01 - 10 ale lunii în curs. Respectarea acestui interval asigură prelucrarea riguroasă și depunerea fără costuri suplimentare a declarațiilor fiscale până la scadența legală de 25 ale lunii.")
    add_a("Art. 12.", "Regimul predărilor întârziate (Majorări pentru prelucrare în regim de urgență)", "În cazul în care Beneficiarul predă documentele cu depășirea termenului standard de 10 ale lunii, pentru compensarea efortului operativ și alocarea de resurse suplimentare de urgență, se pot aplica următoarele ajustări:\n• Predare în intervalul 11 - 18 ale lunii: onorariul lunar curent se majorează cu 15% pentru prelucrarea în regim prioritar, garantându-se depunerea declarațiilor la scadența legală;\n• Predare în intervalul 19 - 23 ale lunii: onorariul lunar curent se majorează cu 30% pentru prelucrare în regim de urgență maximă, depunerea declarațiilor realizându-se la limită.")
    add_a("Art. 13.", "CLAUZĂ CRITICĂ PRIVIND TERMENUL LIMITĂ DE 23 ALE LUNII (EXONERARE DE RĂSPUNDERE)", "În situația în care documentele justificative NU sunt predate sau încărcate complet de către Beneficiar până cel târziu la data de 23 ale lunii, PRESTATORUL NU ÎȘI ASUMĂ OBLIGAȚIA ȘI RESPONSABILITATEA DEPUNERII DECLARAȚIILOR FISCALE AFERENTE ÎN CURSUL LUNII RESPECTIVE (până la termenul legal de 25). În acest caz, documentele vor fi procesate în ordinea normală a fluxului de lucru, iar declarațiile vor fi depuse după termen sau rectificate conform legii. Prestatorul este de plin drept EXONERAT de orice răspundere juridică, fiscală sau materială, iar toate eventualele consecințe negative — incluzând amenzi, dobânzi, penalități de întârziere ANAF, decizii de impunere din oficiu sau popriri — cad în mod exclusiv și irevocabil în sarcina Beneficiarului. În cazul declarațiilor rectificate se va percepe cost suplimentar conform Anexei 2.", bold_red=True)
    add_a("Art. 14.", "Rapoarte și predarea documentelor prelucrate", "Prestatorul pune la dispoziția Beneficiarului, pe platforma securizată de lucru, documentele rezultate în urma prelucrării: balanțele lunare de verificare, jurnalele de cumpărări și vânzări, statele de plată a salariilor, sumele care trebuiesc plătite pentru buget. După încheierea anului fiscal vor fi încărcate declarațiile și confirmările recipiselor de depunere a declarațiilor fiscale.")

    doc.add_heading("CAPITOLUL VI. DREPTURILE ȘI OBLIGAȚIILE PĂRȚILOR", level=1)
    add_a("Art. 15.", "Obligațiile esențiale ale Prestatorului", "a) Să asigure înregistrarea cronologică și sistematică a documentelor în partidă dublă; b) Să întocmească jurnalele contabile, balanțele lunare de verificare și registrele obligatorii conform legii; c) Să calculeze obligațiile fiscale și să transmită declarațiile fiscale în termenele legale, sub condiția primirii documentelor complete în termenele contractuale; d) Să asigure evidența salariaților și operarea instrucțiunilor primite în REGES-ONLINE conform Anexei nr. 3; e) Să informeze Beneficiarul asupra modificărilor legislative relevante cu impact asupra activității sale; f) Să asigure confidențialitatea strictă a datelor și informațiilor primite.")
    add_a("Art. 16.", "Obligațiile esențiale ale Beneficiarului", "a) Să pună la dispoziția Prestatorului documente primare complete, corecte, lizibile și legale, exclusiv aferente activității economice a societății; b) Să respecte termenele de predare a documentelor (01 - 10 ale lunii); c) Să achite onorariul convenit și contravaloarea serviciilor suplimentare la termenele scadente; d) Să asigure accesul Prestatorului în SPV (Spațiul Privat Virtual) și să transmită extrasele bancare complete în format electronic; e) Să comunice în scris și în timp util orice modificare privind salariații (angajări, încetări, concedii medicale etc.); f) Să furnizeze orice lămuriri suplimentare solicitate de Prestator pentru reflectarea fidelă a tranzacțiilor.")

    doc.add_heading("CAPITOLUL VII. DELIMITAREA RĂSPUNDERII ȘI LIMITAREA DESPĂGUBIRILOR", level=1)
    add_a("Art. 17.", "Delimitarea răspunderii patrimoniale", "Prestatorul răspunde exclusiv de reflectarea corectă în contabilitate a operațiunilor economice și de calculul corect al taxelor, pe baza documentelor justificative puse la dispoziție de Beneficiar. Prestatorul se obligă să suporte amenzile și penalitățile aplicate de organele fiscale rezultate exclusiv din culpa sa profesională dovedită. Prestatorul NU răspunde în niciun fel de: a) Gestiunea fizică a patrimoniului, recepția bunurilor, lipsurile din inventar sau efectuarea inventarierilor anuale; b) Legalitatea, autenticitatea și oportunitatea economică a documentelor emise de partenerii comerciali ai Beneficiarului; c) Emiterea facturilor proprii și transmiterea lor în sistemul național RO e-Factura; d) Deducerea integrală (100%) a cheltuielilor și TVA auto în lipsa foilor de parcurs complete și corecte puse la dispoziție de Beneficiar; e) Calculul și plata taxelor și impozitelor locale (clădiri, autoturisme, terenuri); f) Modul în care a fost ținută contabilitatea anterior preluării mandatului de către Prestator sau de întârzierile cauzate de refacerea evidențelor anterioare.")
    add_a("Art. 18.", "Plafonarea despăgubirilor", "Răspunderea profesională totală a Prestatorului pentru eventuale prejudicii materiale directe cauzate Beneficiarului exclusiv din culpa sa profesională dovedită este expres și irevocabil plafonată la maximum contravaloarea a 3 (trei) onorarii lunare contractuale de contabilitate.")
    add_a("Art. 19.", "Dreptul de retenție asupra bazei de date electronice", "Prestatorul își rezervă dreptul legitim de retenție asupra bazei de date contabile generate în softul de specialitate (ex: Saga) și asupra dosarelor contabile până la achitarea integrală de către Beneficiar a tuturor onorariilor și debitelor restante. În lipsa plății integrale, baza de date contabilă se conservă pentru maximum 6 luni de la încetarea contractului, după care poate fi ștearsă din sistemele Prestatorului fără altă notificare prealabilă.")
    add_a("Art. 20.", "Clauza de non-racolare a personalului", "Beneficiarul se obligă ferm să nu recruteze, să nu angajeze și să nu contracteze direct sau indirect servicii cu personalul angajat sau colaboratorii de specialitate ai Prestatorului, pe durata derulării prezentului contract și pe o perioadă de 12 luni de la încetarea acestuia. În cazul încălcării acestei obligații, Beneficiarul va achita Prestatorului o despăgubire forfetară fermă egală cu contravaloarea a 24 de luni de onorariu contractual de contabilitate.")

    doc.add_heading("CAPITOLUL VIII. PROTECȚIA DATELOR CU CARACTER PERSONAL (RGPD)", level=1)
    add_a("Art. 21.", "Cadrul general", "Părțile se obligă să respecte strict normele Regulamentului (UE) 2016/679 (RGPD) și a Legii nr. 190/2018.")
    add_a("Art. 22.", "Rolurile părților", "Beneficiarul = Operator de date; Prestatorul = Persoană Împuternicită de operator (Împuternicit), prelucrând datele exclusiv pe baza instrucțiunilor Beneficiarului.")
    add_a("Art. 23.", "Categorii de date", "Salariați, administratori, asociați, parteneri comerciali (nume, CNP, acte identitate, pontaje, salarii, concedii medicale).")
    add_a("Art. 24.", "Obligațiile Operatorului", "Deține temeiul legal, informează persoanele vizate, emite instrucțiuni conforme.")
    add_a("Art. 25.", "Obligațiile Împuternicitului", "Prelucrează strict pe bază de instrucțiuni, asigură confidențialitatea și securitatea IT, notifică breșele în max. 48 ore.")
    add_a("Art. 26.", "Sub-împuterniciți", "Autorizare generală pentru Saga Software, Google Workspace/Drive, servicii de semnătură electronică calificată.")
    add_a("Art. 27.", "Încetarea contractului", "Predarea bazei de date după achitarea debitelor. Conservare arhivă conform Legii contabilității nr. 82/1991.")

    doc.add_heading("CAPITOLUL IX. CONFIDENȚIALITATEA COMERCIALĂ", level=1)
    add_a("Art. 28.", "Clauza de confidențialitate", "Toate informațiile comerciale și contabile sunt strict confidențiale pe toată durata contractului și 2 ani după încetare.")

    doc.add_heading("CAPITOLUL X. FORȚA MAJORĂ, ÎNCETAREA CONTRACTULUI ȘI LITIGII", level=1)
    add_a("Art. 29.", "Forța majoră", "Apără de răspundere cu notificare în 5 zile. Încetare după 30 de zile fără daune.")
    add_a("Art. 30.", "Încetarea contractului", "Acord scris; denunțare cu preaviz 30 de zile; reziliere de drept la neplată >30 zile; insolvență/faliment.")
    add_a("Art. 31.", "Soluționarea litigiilor", "Conciliere/arbitraj la Comisia de Arbitraj a CECCAR; instanțele competente de la sediul Prestatorului.")

    doc.add_heading("CAPITOLUL XI. DATE DE CONTACT ȘI COMUNICARE OPERATIVĂ", level=1)
    add_a("Art. 32.", "Comunicări", f"E-mail Prestator: {prest['email']} (Tel: {prest['telefon']}) | E-mail Beneficiar: {client_data.get('email', '____________________')} (Tel: {client_data.get('telefon', '__________')})")

    # Semnaturi
    p_sf = doc.add_paragraph()
    set_p(p_sf, 4, 2)
    p_sf.add_run(f"Încheiat și semnat astăzi, {data_ctr}, în 2 (două) exemplare originale cu valoare juridică egală.")
    
    ts = doc.add_table(rows=1, cols=2)
    ts.alignment = WD_TABLE_ALIGNMENT.CENTER
    ts.autofit = False
    ts.rows[0].cells[0].width = Inches(3.5)
    ts.rows[0].cells[0].width = Inches(3.5)
    ts.rows[0].cells[0].paragraphs[0].add_run(f"PRESTATOR:\n{prest['nume']}\nPrin Administrator: {prest['reprezentant']}\n\nSemnătura: ______________\nȘtampila:")
    ts.rows[0].cells[0].paragraphs[0].add_run(f"BENEFICIAR:\n{client_data.get('nume', 'CLIENT SRL')}\nPrin Administrator: {client_data.get('reprezentant', 'ADMIN')}\n\nSemnătura: ______________\nȘtampila:")

    # Anexa 1
    doc.add_page_break()
    doc.add_heading("ANEXA NR. 1: PRESTAȚII DE BAZĂ (PACHET SERVICII CONTABILITATE ȘI SALARIZARE)", level=1)
    p_a1 = doc.add_paragraph()
    set_p(p_a1, 2, 4)
    p_a1.add_run(
        "A. Evidență Financiar-Contabilă și Registre Obligatorii:\n"
        "1. Contabilizarea documentelor în partidă dublă (intrări, ieșiri, bancă, casă, deconturi).\n"
        "2. Evidența analitică/sintetică clienți, furnizori, trezorerie, imobilizări.\n"
        "3. Calcul amortizare lunară a mijloacelor fixe.\n"
        "4. Balanțe lunare de verificare (patru serii de egalități).\n"
        "5. Registrul Jurnal, Registrul Inventar, Jurnal Cumpărări, Jurnal Vânzări, Registru Evidență Fiscală.\n"
        "6. Fișe de cont analitice și Carte Mare (Maestru Șah).\n\n"
        "B. Salarizare și Resurse Umane (Pachet de bază):\n"
        "1. State de plată lunare, pontaje, fluturași pe baza datelor primite.\n"
        "2. Calcul rețineri salariale, CAS, CASS, CAM, impozit pe salarii.\n"
        "3. Calcul concedii medicale și impozit tichete de masă/vacanță/cadou.\n"
        "4. Ordine de plată pentru taxele salariale la buget.\n"
        "5. Adeverințe tip medic/spital.\n\n"
        "C. Declarații Fiscale Periodice:\n"
        "1. Calcul impozit micro/profit, dividende, rețineri la sursă.\n"
        "2. Decont TVA (D300/301).\n"
        "3. Declarație informativă livrări/achiziții interne (D394).\n"
        "4. Declarație recapitulativă livrări/achiziții intracomunitare (D390 VIES).\n"
        "5. Declarația privind obligațiile la buget (D100).\n"
        "6. Declarația privind contribuțiile sociale și asigurații (D112).\n"
        "7. Fișierul Standard de Control Fiscal (SAF-T / D406), conform legii (fără stocuri/producție).\n"
        "8. Declarația de mențiuni vector fiscal (Formularul 700)."
    )

    # Anexa 2
    doc.add_page_break()
    doc.add_heading("ANEXA NR. 2: GRILA OFICIALĂ DE PREȚURI ȘI TARIFE SUPLIMENTARE (2026)", level=1)
    doc.add_heading("I. Pentru Activități de PRESTĂRI SERVICII", level=2)
    hdr_s_w = ["Nr. tranzacții", "0 - 40", "40 - 100", "101 - 200", "201 - 300", "301 - 400", "401 - 500"]
    rows_s_w = [
        ["Microîntreprindere, fără TVA", "500 lei", "600 lei", "700 lei", "800 lei", "900 lei", "1.000 lei"],
        ["Cu TVA (D300 și D394)", "600 lei", "700 lei", "800 lei", "900 lei", "1.000 lei", "1.100 lei"],
        ["Cu D390 (servicii intrastat)", "650 lei", "750 lei", "850 lei", "950 lei", "1.050 lei", "1.150 lei"],
        ["Impozit pe profit", "700 lei", "800 lei", "900 lei", "1.000 lei", "1.100 lei", "1.200 lei"],
        ["Consultanță inclusă", "15 min", "15 min", "15 min", "15 min", "30 min", "30 min"]
    ]
    t1 = doc.add_table(rows=len(rows_s_w)+1, cols=7)
    t1.alignment = WD_TABLE_ALIGNMENT.CENTER
    t1.autofit = False
    for r_i, rd in enumerate([hdr_s_w] + rows_s_w):
        for c_i, val in enumerate(rd):
            c = t1.rows[r_i].cells[c_i]
            c.width = Inches(1.8 if c_i == 0 else 0.85)
            set_cell_margins(c, 40, 40, 50, 50)
            p = c.paragraphs[0]
            if r_i == 0:
                set_cell_background(c, "1B365D")
                r = p.add_run(val)
                r.font.bold = True
                r.font.color.rgb = RGBColor(255, 255, 255)
            else:
                p.add_run(val)

    doc.add_heading("II. Pentru Activități de MARFĂ / GESTIUNE", level=2)
    hdr_m_w = ["Nr. tranzacții", "6 - 20", "21 - 100", "101 - 200", "201 - 300", "301 - 400", "401 - 500"]
    rows_m_w = [
        ["Microîntreprindere, fără TVA", "700 lei", "800 lei", "900 lei", "1.000 lei", "1.100 lei", "1.200 lei"],
        ["Cu TVA (D300 și D394)", "1.000 lei", "1.100 lei", "1.200 lei", "1.300 lei", "1.400 lei", "1.500 lei"],
        ["Cu D390 (servicii intrastat)", "1.050 lei", "1.150 lei", "1.250 lei", "1.350 lei", "1.450 lei", "1.550 lei"],
        ["Impozit pe profit", "1.150 lei", "1.250 lei", "1.350 lei", "1.450 lei", "1.550 lei", "1.650 lei"],
        ["Consultanță inclusă", "15 min", "15 min", "15 min", "15 min", "30 min", "30 min"]
    ]
    t2 = doc.add_table(rows=len(rows_m_w)+1, cols=7)
    t2.alignment = WD_TABLE_ALIGNMENT.CENTER
    t2.autofit = False
    for r_i, rd in enumerate([hdr_m_w] + rows_m_w):
        for c_i, val in enumerate(rd):
            c = t2.rows[r_i].cells[c_i]
            c.width = Inches(1.8 if c_i == 0 else 0.85)
            set_cell_margins(c, 40, 40, 50, 50)
            p = c.paragraphs[0]
            if r_i == 0:
                set_cell_background(c, "2B4C7E")
                r = p.add_run(val)
                r.font.bold = True
                r.font.color.rgb = RGBColor(255, 255, 255)
            else:
                p.add_run(val)

    doc.add_heading("III. Tarife Suplimentare (Activități la cerere 2026)", level=2)
    suplim_w = [
        ("De la a 3-a bancă / cont bancar", "50 lei/bancă"),
        ("De la al 2-lea contract de leasing", "10 lei/lună"),
        ("Deschidere Revisal - primul contract", "200 lei"),
        ("Calcul salariu, stat de plată, pontaj, D112, OP-uri", "50 lei/salariat/lună"),
        ("Statistica", "începând de la 350 lei"),
        ("Obținere cod TVA la înființare / intracomunitar", "300 lei"),
        ("Obținere număr EORI (vamă)", "300 lei"),
        ("Dosar recuperare indemnizații FNUASS", "200 lei"),
        ("Obținere NIF / Certificat rezidență fiscală", "200 lei / 300 lei"),
        ("Depunere declarație casă de marcat", "50 lei"),
        ("Dosar depășire plafon de scutire TVA", "300 lei"),
        ("Înregistrare punct de lucru ANAF (fără / cu salariați)", "100 lei"),
        ("Întocmire Regulament Intern", "începând de la 350 lei"),
        ("Achiziționare Registru Unic de Control (RUC)", "începând de la 100 lei (+ cost registru)"),
        ("Asistență la rambursare TVA", "începând de la 500 lei"),
        ("Reactivare firmă / Activare cod TVA anulat", "300 lei"),
        ("Completare acte solicitate de bănci", "începând de la 300 lei"),
        ("Certificat constatator ONRC", "130 lei"),
        ("Consultanță fiscală suplimentară", "50 EURO/oră"),
        ("Secretariat - aranjare documente (până la 50 doc.)", "50 lei"),
        ("Bilanț interimar / Situații financiare anuale", "costul unei luni (minim 350 lei)"),
        ("Întocmire manual de politici și proceduri contabile", "500 lei")
    ]
    t3 = doc.add_table(rows=len(suplim_w)+1, cols=2)
    t3.alignment = WD_TABLE_ALIGNMENT.CENTER
    t3.autofit = False
    for r_i, (s_name, s_pret) in enumerate([("Serviciu", "Tarif")] + suplim_w):
        for c_i, val in enumerate([s_name, s_pret]):
            c = t3.rows[r_i].cells[c_i]
            c.width = Inches(4.8 if c_i == 0 else 2.2)
            set_cell_margins(c, 30, 30, 50, 50)
            p = c.paragraphs[0]
            if r_i == 0:
                set_cell_background(c, "1B365D")
                r = p.add_run(val)
                r.font.bold = True
                r.font.color.rgb = RGBColor(255, 255, 255)
            else:
                p.add_run(val)

    # Anexa 3
    doc.add_page_break()
    doc.add_heading("ANEXA NR. 3: PROCEDURA PRIVIND ACCESUL ȘI TRANSMITEREA ÎN REGES-ONLINE ȘI PROTECȚIA DATELOR", level=1)
    add_a("Art. 1.", "Cadrul legal și crearea contului", "În temeiul Procedurii din 01.08.2025 a Ministerului Muncii, Prestatorul prin administrator va face demersurile pentru crearea și gestionarea contului de acces al angajatorului.")
    add_a("Art. 2.", "Persoana desemnată și RGPD", "Reprezentantul desemnat al Prestatorului este persoana împuternicită prin decizie scrisă conform Regulamentului (UE) 2016/679 pentru operarea datelor în Registru.")
    add_a("Art. 3.", "Instrucțiuni scrise", "Beneficiarul va transmite instrucțiuni scrise prin e-mail pentru orice contract de muncă, modificare, suspendare sau încetare.")
    add_a("Art. 4.", "Termene legale și răspundere", "Obligația transmiterii datelor către Prestator anterior termenelor legale revine exclusiv Beneficiarului (angajare cu cel puțin 1 zi înainte). Prestatorul răspunde de operarea la timp a instrucțiunilor primite.")
    add_a("Art. 5.", "Salarii asigurate de Beneficiar", "Dacă Beneficiarul calculează intern salariile, are obligația de a înștiința neîntârziat Prestatorul pentru operarea notelor contabile.")

    ts3 = doc.add_table(rows=1, cols=2)
    ts3.alignment = WD_TABLE_ALIGNMENT.CENTER
    ts3.autofit = False
    ts3.rows[0].cells[0].width = Inches(3.5)
    ts3.rows[0].cells[0].width = Inches(3.5)
    ts3.rows[0].cells[0].paragraphs[0].add_run(f"PRESTATOR:\n{prest['nume']}\nPrin Administrator: {prest['reprezentant']}\n\nSemnătura: ______________\nȘtampila:")
    ts3.rows[0].cells[0].paragraphs[0].add_run(f"BENEFICIAR:\n{client_data.get('nume', 'CLIENT SRL')}\nPrin Administrator: {client_data.get('reprezentant', 'ADMIN')}\n\nSemnătura: ______________\nȘtampila:")

    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

print("Builder module fully updated with PDF and DOCX generators!")

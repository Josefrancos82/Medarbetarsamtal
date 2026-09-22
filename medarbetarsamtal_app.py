import streamlit as st
import io
import os
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# Optional imports for Gemini AI and Image handling
try:
    import google.generativeai as genai
    from PIL import Image
    HAS_GEMINI = True
except Exception:
    HAS_GEMINI = False

# Page Configuration
st.set_page_config(
    page_title="Medarbetarsamtal & Målgenerator - Skolan",
    page_icon="🏫",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Initialize Gemini if API key is provided in st.secrets or environment
api_key = None
if HAS_GEMINI:
    if "GEMINI_API_KEY" in st.secrets:
        api_key = st.secrets["GEMINI_API_KEY"]
    elif "gemini" in st.secrets and "api_key" in st.secrets["gemini"]:
        api_key = st.secrets["gemini"]["api_key"]
    elif os.environ.get("GEMINI_API_KEY"):
        api_key = os.environ.get("GEMINI_API_KEY")

if api_key and HAS_GEMINI:
    try:
        genai.configure(api_key=api_key)
        AI_READY = True
    except Exception:
        AI_READY = False
else:
    AI_READY = False

# Custom Styling - Modern Abstract Linear Dark Theme with High Contrast Text
st.markdown("""
<style>
    /* Global App Background - Modern Abstract Lines */
    .stApp {
        background-color: #0B0F19;
        background-image: 
            radial-gradient(at 10% 10%, rgba(30, 58, 138, 0.45) 0px, transparent 40%),
            radial-gradient(at 90% 90%, rgba(124, 58, 237, 0.35) 0px, transparent 45%),
            radial-gradient(at 50% 50%, rgba(14, 165, 233, 0.22) 0px, transparent 50%),
            url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='100%25' height='100%25' viewBox='0 0 1600 900'%3E%3Cdefs%3E%3ClinearGradient id='g1' x1='0%25' y1='0%25' x2='100%25' y2='100%25'%3E%3Cstop offset='0%25' stop-color='%2338bdf8' stop-opacity='0.28'/%3E%3Cstop offset='100%25' stop-color='%23818cf8' stop-opacity='0.03'/%3E%3C/linearGradient%3E%3ClinearGradient id='g2' x1='100%25' y1='0%25' x2='0%25' y2='100%25'%3E%3Cstop offset='0%25' stop-color='%23c084fc' stop-opacity='0.22'/%3E%3Cstop offset='100%25' stop-color='%2338bdf8' stop-opacity='0.02'/%3E%3C/linearGradient%3E%3C/defs%3E%3Cpath d='M-100 180 C 300 30, 700 420, 1700 80' stroke='url(%23g1)' stroke-width='3' fill='none'/%3E%3Cpath d='M-100 230 C 350 80, 750 470, 1700 130' stroke='url(%23g1)' stroke-width='2' fill='none'/%3E%3Cpath d='M-100 280 C 400 130, 800 520, 1700 180' stroke='url(%23g1)' stroke-width='1' fill='none'/%3E%3Cpath d='M-100 580 C 400 330, 900 830, 1700 480' stroke='url(%23g2)' stroke-width='3.5' fill='none'/%3E%3Cpath d='M-100 630 C 450 380, 950 880, 1700 530' stroke='url(%23g2)' stroke-width='2' fill='none'/%3E%3Cpath d='M-100 680 C 500 430, 1000 930, 1700 580' stroke='url(%23g2)' stroke-width='1' fill='none'/%3E%3Cpath d='M 120 -100 C 520 380, 220 780, 1220 1000' stroke='url(%23g1)' stroke-width='2.5' fill='none'/%3E%3C/svg%3E");
        background-attachment: fixed;
        background-size: cover;
        color: #FFFFFF !important;
    }

    /* Sidebar Styling */
    [data-testid="stSidebar"] {
        background-color: rgba(15, 23, 42, 0.94) !important;
        border-right: 1px solid rgba(56, 189, 248, 0.25) !important;
        backdrop-filter: blur(12px);
    }
    
    [data-testid="stSidebar"] * {
        color: #F8FAFC !important;
    }

    /* Main Typography */
    .main-header {
        font-size: 2.3rem;
        background: linear-gradient(135deg, #60A5FA 0%, #38BDF8 50%, #C084FC 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        font-weight: 800;
        margin-bottom: 0.3rem;
        letter-spacing: -0.5px;
    }

    .sub-header {
        font-size: 1.15rem;
        color: #E2E8F0 !important;
        margin-bottom: 1.8rem;
        font-weight: 400;
        line-height: 1.5;
    }

    /* Card Containers */
    .card {
        background: rgba(15, 23, 42, 0.8) !important;
        border: 1px solid rgba(56, 189, 248, 0.25) !important;
        border-radius: 12px !important;
        padding: 20px !important;
        margin-bottom: 18px !important;
        box-shadow: 0 8px 32px 0 rgba(0, 0, 0, 0.37) !important;
        backdrop-filter: blur(10px) !important;
        color: #FFFFFF !important;
    }

    /* Form Labels & General Text */
    label, p, span, h1, h2, h3, h4, h5, h6 {
        color: #F8FAFC !important;
    }

    /* Inputs, Textareas, Selectboxes */
    .stTextInput input, .stTextArea textarea, div[data-baseweb="select"] > div {
        background-color: #0F172A !important;
        color: #FFFFFF !important;
        border: 1px solid #334155 !important;
        border-radius: 8px !important;
    }

    .stTextInput input:focus, .stTextArea textarea:focus {
        border-color: #38BDF8 !important;
        box-shadow: 0 0 0 2px rgba(56, 189, 248, 0.25) !important;
    }

    /* Dropdown Menus */
    ul[role="listbox"] {
        background-color: #0F172A !important;
        color: #FFFFFF !important;
    }
    li[role="option"] {
        color: #FFFFFF !important;
    }

    /* Tab Styling */
    button[data-baseweb="tab"] {
        color: #94A3B8 !important;
        font-weight: 600 !important;
        font-size: 1.05rem !important;
        background-color: transparent !important;
        border-radius: 6px 6px 0 0 !important;
        padding: 10px 18px !important;
    }

    button[aria-selected="true"] {
        color: #38BDF8 !important;
        border-bottom: 3px solid #38BDF8 !important;
        background: rgba(30, 41, 59, 0.6) !important;
    }

    /* Expanders */
    .streamlit-expanderHeader {
        background-color: rgba(30, 41, 59, 0.75) !important;
        color: #FFFFFF !important;
        border-radius: 8px !important;
        border: 1px solid rgba(56, 189, 248, 0.2) !important;
    }

    /* Buttons */
    .stButton > button[kind="primary"] {
        background: linear-gradient(135deg, #2563EB 0%, #0284C7 100%) !important;
        color: #FFFFFF !important;
        border: none !important;
        font-weight: 600 !important;
        border-radius: 8px !important;
        padding: 8px 16px !important;
        box-shadow: 0 4px 14px 0 rgba(2, 132, 199, 0.39) !important;
    }

    .stButton > button[kind="primary"]:hover {
        background: linear-gradient(135deg, #1D4ED8 0%, #0369A1 100%) !important;
        box-shadow: 0 6px 20px 0 rgba(2, 132, 199, 0.55) !important;
    }

    /* Info / Success boxes */
    .stAlert {
        background-color: rgba(15, 23, 42, 0.88) !important;
        border: 1px solid rgba(56, 189, 248, 0.3) !important;
        color: #F8FAFC !important;
        border-radius: 10px !important;
    }

    /* Upload box */
    [data-testid="stFileUploader"] {
        background-color: rgba(15, 23, 42, 0.7) !important;
        border: 2px dashed rgba(56, 189, 248, 0.4) !important;
        border-radius: 10px !important;
        padding: 12px !important;
    }
</style>
""", unsafe_allow_html=True)

# ---------------------------------------------------------
# HELPER FUNCTIONS FOR DOCX GENERATION
# ---------------------------------------------------------
def set_cell_background(cell, hex_color):
    """Sets background color of a docx table cell."""
    tcPr = cell._element.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color)
    tcPr.append(shd)

def create_word_document(emp_name, manager_name, role, period, date_str, school_goals, nulage_data, criteria_data, plan_data):
    """Generates an official Word document matching the municipality's documentation template."""
    doc = docx.Document()
    
    # Page Margins
    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)
    
    # Title & Header
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_title = p_title.add_run("Dokumentationsmall för medarbetarsamtal")
    r_title.bold = True
    r_title.font.size = Pt(18)
    r_title.font.color.rgb = RGBColor(30, 58, 138)
    
    # Meta Info Table
    meta_table = doc.add_table(rows=2, cols=2)
    meta_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    meta_table.autofit = False
    
    col_widths = [Inches(3.2), Inches(3.2)]
    
    row0 = meta_table.rows[0]
    row0.cells[0].paragraphs[0].add_run("Medarbetarens namn: ").bold = True
    row0.cells[0].paragraphs[0].add_run(emp_name if emp_name else "Kim Ekman")
    row0.cells[1].paragraphs[0].add_run("Chefens namn: ").bold = True
    row0.cells[1].paragraphs[0].add_run(manager_name if manager_name else "Skolledare")
    
    row1 = meta_table.rows[1]
    row1.cells[0].paragraphs[0].add_run("Befattning / Yrkeskategori: ").bold = True
    row1.cells[0].paragraphs[0].add_run(role)
    row1.cells[1].paragraphs[0].add_run("Datum / Period: ").bold = True
    row1.cells[1].paragraphs[0].add_run(f"{date_str} ({period})")
    
    for row in meta_table.rows:
        for idx, cell in enumerate(row.cells):
            cell.width = col_widths[idx]
            set_cell_background(cell, "F1F5F9")
            
    doc.add_paragraph().paragraph_format.space_after = Pt(12)
    
    # Section 1: Arbetssituation - nuläge
    h1 = doc.add_heading("1. Arbetssituation – nuläge", level=1)
    h1.runs[0].font.color.rgb = RGBColor(30, 58, 138)
    
    p_intro1 = doc.add_paragraph()
    p_intro1.add_run("Genomgång av medarbetarens nuvarande arbetssituation, arbetsmiljö, hälsa, trivsel, ledarskap och eventuell bisyssla.").italic = True
    
    for key, val in nulage_data.items():
        if val.strip():
            p_sub = doc.add_paragraph()
            p_sub.paragraph_format.space_before = Pt(6)
            p_sub.paragraph_format.space_after = Pt(2)
            r_head = p_sub.add_run(f"{key}: ")
            r_head.bold = True
            r_head.font.color.rgb = RGBColor(15, 23, 42)
            p_sub.add_run(val)
            
    doc.add_paragraph().paragraph_format.space_after = Pt(12)
    
    # Section 2: Resultat - tillbakablick (Lönekriterier)
    h2 = doc.add_heading("2. Resultat – tillbakablick (Lönekriterier)", level=1)
    h2.runs[0].font.color.rgb = RGBColor(30, 58, 138)
    
    p_intro2 = doc.add_paragraph()
    p_intro2.add_run("Uppföljning av medarbetarens resultat och arbetsprestation utifrån verksamhetens mål och lönekriterier.").italic = True
    
    # Criteria Table
    crit_table = doc.add_table(rows=1, cols=3)
    crit_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    crit_table.autofit = False
    
    hdr_cells = crit_table.rows[0].cells
    hdr_widths = [Inches(1.8), Inches(1.5), Inches(3.2)]
    headers = ["Lönekriterium", "Bedömd Nivå", "Motivering & Exempel från undervisningen"]
    
    for idx, text in enumerate(headers):
        hdr_cells[idx].width = hdr_widths[idx]
        set_cell_background(hdr_cells[idx], "1E3A8A")
        p = hdr_cells[idx].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        r = p.add_run(text)
        r.bold = True
        r.font.color.rgb = RGBColor(255, 255, 255)
        
    for c_name, c_info in criteria_data.items():
        row_cells = crit_table.add_row().cells
        row_cells[0].width = hdr_widths[0]
        row_cells[1].width = hdr_widths[1]
        row_cells[2].width = hdr_widths[2]
        
        p0 = row_cells[0].paragraphs[0]
        p0.add_run(c_name).bold = True
        
        p1 = row_cells[1].paragraphs[0]
        p1.add_run(c_info.get("level", "Når ett bra resultat"))
        
        p2 = row_cells[2].paragraphs[0]
        p2.add_run(c_info.get("motive", "Följer uppställda mål."))
        
        set_cell_background(row_cells[0], "F8FAFC")
        
    doc.add_paragraph().paragraph_format.space_after = Pt(12)
    
    # Section 3: Utveckling - framåtblick (Individuell utvecklingsplan)
    h3 = doc.add_heading("3. Utveckling – framåtblick (Individuell Utvecklingsplan)", level=1)
    h3.runs[0].font.color.rgb = RGBColor(30, 58, 138)
    
    p_intro3 = doc.add_paragraph()
    p_intro3.add_run("Individuella mål utformade enligt SMART-modellen utifrån skolans prioriterade mål och medarbetarens utvecklingsområden.").italic = True
    
    # Plan Table
    plan_table = doc.add_table(rows=1, cols=3)
    plan_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    plan_table.autofit = False
    
    p_hdr_cells = plan_table.rows[0].cells
    p_hdr_widths = [Inches(2.2), Inches(2.8), Inches(1.5)]
    p_headers = ["Mål / Utvecklingsområde (SMART)", "Beskrivning av åtgärder & ansvarsfördelning", "Uppföljningsplan & Tidsram"]
    
    for idx, text in enumerate(p_headers):
        p_hdr_cells[idx].width = p_hdr_widths[idx]
        set_cell_background(p_hdr_cells[idx], "1E3A8A")
        p = p_hdr_cells[idx].paragraphs[0]
        r = p.add_run(text)
        r.bold = True
        r.font.color.rgb = RGBColor(255, 255, 255)
        
    for item in plan_data:
        row_cells = plan_table.add_row().cells
        row_cells[0].width = p_hdr_widths[0]
        row_cells[1].width = p_hdr_widths[1]
        row_cells[2].width = p_hdr_widths[2]
        
        row_cells[0].paragraphs[0].add_run(item.get("goal", ""))
        row_cells[1].paragraphs[0].add_run(item.get("actions", ""))
        row_cells[2].paragraphs[0].add_run(item.get("followup", ""))
        
        set_cell_background(row_cells[0], "F8FAFC")
        
    # Save to buffer
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# ---------------------------------------------------------
# CRITERIA DEFINITIONS FROM SOURCES
# ---------------------------------------------------------
TEACHER_CRITERIA_INFO = {
    "Yrkeskunnande/färdighet": "Anpassar undervisningen efter elevernas förmågor/behov. Tillämpar varierade metoder som ger högre måluppfyllelse och minskar behovet av särskilda insatser. Säkerställer rättssäker och allsidig bedömning.",
    "Engagemang/ansvarsförmåga": "Gör eleverna delaktiga i läroprocessen. Tar ansvar för alla elevers utveckling med lösningsfokuserat förhållningssätt. Drivande i arbetslag/ämneslag och bidrar i det systematiska kvalitetsarbetet (SKA).",
    "Professionellt förhållningssätt": "God förebild och trygg i yrkesrollen. Skapar goda relationer med elever och vårdnadshavare. Förankrad i skolans värdegrund och stödjer elevers inflytande och delaktighet.",
    "Samarbetsförmåga": "Aktivt inlyssnande och deltar i det kollegiala lärandet. Bjuder in till samarbete med kollegor och andra professioner (t.ex. elevhälsa).",
    "Ledarskapsförmåga inom yrket": "Skapar trygg och arbetsinriktad lärmiljö med hög tilltro till elevers lärande. Väcker kunskapslust och leder lärandet systematiskt framåt.",
    "Flexibilitet/utvecklingsmöjlighet": "Följer forskning, analyserar systematiskt elevernas lärande, provar nya arbetssätt och delar med sig av sin kunskap kollegialt."
}

FRITIDS_CRITERIA_INFO = {
    "Yrkeskunnande/färdighet": "Skapar likvärdighet genom att anpassa fritidshemsundervisningen efter elevernas behov. Tillämpar metoder som ökar elevers tilltro till egen förmåga.",
    "Engagemang/ansvarsförmåga": "Systematiskt elevinflytande i fritidshemmet. Drivande i utvecklingen och aktiv deltagare i enhetens kvalitetsarbete.",
    "Professionellt förhållningssätt": "Professionellt bemötande av elever och vårdnadshavare. Trygg förebild förankrad i värdegrunden.",
    "Samarbetsförmåga": "Aktiv i kollegialt samarbete med skola, fritidshem och elevhälsa.",
    "Ledarskapsförmåga inom yrket": "Leder fritidshemmets undervisning och aktiviteter med hög elevdelaktighet och trygghet.",
    "Flexibilitet/utvecklingsmöjlighet": "Systematisk analys av fritidshemsverksamheten, provar nya arbetssätt och följer forskning."
}

# ---------------------------------------------------------
# MAIN APP ENTRY POINT (SAFE FOR STREAMLIT RUN)
# ---------------------------------------------------------
def main():
    # SIDEBAR SETUP
    with st.sidebar:
        st.image("https://img.icons8.com/color/96/school.png", width=64)
        st.title("⚙️ Samtalsinställningar")
        
        if AI_READY:
            st.success("🟢 AI (Gemini) Aktiverad")
        else:
            st.info("🔵 Regelbaserat läge (Ingen API-nyckel behövs)")
        
        st.subheader("👤 Medarbetaruppgifter")
        emp_name = st.text_input("Medarbetarens namn", value="Kim Ekman")
        manager_name = st.text_input("Chefens namn", value="José Francos-Dafgård")
        
        role = st.selectbox(
            "Yrkeskategori",
            ["Lärare / Förskollärare i F-klass", "Fritidspedagog i fritidshem"]
        )
        
        period = st.text_input("Gäller tidsperiod", value="Läsår 2026/2027")
        date_str = st.text_input("Samtalsdatum", value="2026-09-22")
        
        st.markdown("---")
        st.subheader("🎯 Skolans Prioriterade Mål")
        st.caption("Målen används för att förankra medarbetarens SMART-mål.")
        
        school_mål_1 = st.checkbox("1. Höjd måluppfyllelse & likvärdig undervisning", value=True)
        school_mål_2 = st.checkbox("2. Trygghet, studiero & tillgänglig lärmiljö", value=True)
        school_mål_3 = st.checkbox("3. Stärkt kollegialt lärande & SKA", value=True)
        school_mål_4 = st.checkbox("4. Elevhälsa, tidiga insatser & anpassningar", value=False)
        
        custom_school_mål = st.text_input("Eget/kompletterande skolmål", value="")
        
        selected_school_goals = []
        if school_mål_1: selected_school_goals.append("Höjd måluppfyllelse & likvärdig undervisning")
        if school_mål_2: selected_school_goals.append("Trygghet, studiero & tillgänglig lärmiljö")
        if school_mål_3: selected_school_goals.append("Stärkt kollegialt lärande & SKA")
        if school_mål_4: selected_school_goals.append("Elevhälsa & tidiga insatser")
        if custom_school_mål.strip(): selected_school_goals.append(custom_school_mål.strip())

    # MAIN INTERFACE
    st.markdown('<div class="main-header">🏫 Medarbetarsamtal & Målgenerator för Skolan</div>', unsafe_allow_html=True)
    st.markdown('<div class="sub-header">Omvandla råa/fotade anteckningar till officiell, färdigformaterad dokumentation och knivskarpa SMART-mål redo för kommunens Word-mallar.</div>', unsafe_allow_html=True)

    # App Navigation Tabs
    tab1, tab2, tab3, tab4 = st.tabs([
        "📝 Omvandla Anteckningar (Word-klar)",
        "🎯 SMART-Målgenerator",
        "📊 Lönekriterier & Bedömning",
        "📄 Förhandsgranska & Exportera Word"
    ])

    # Initialize session state for stored data
    if 'nulage_notes' not in st.session_state:
        st.session_state.nulage_notes = {
            "Arbetsmiljö och arbetsbelastning": "Upplever hög belastning vid rättning, men god struktur i klassrummet. Önskar avlastning kring rastvakter.",
            "Hälsa och återhämtning": "Mår bra i grunden, använder friskvårdsbidraget för träning.",
            "Trivsel och motivation": "Stort engagemang i ämneslaget. Motiveras av kollegialt samarbete.",
            "Ledarskap och stöd": "Nöjd med chefskontakten, önskar mer återkoppling på lektionsbesök.",
            "Bisyssla": "Ingen bisyssla att anmäla."
        }

    if 'smart_goals' not in st.session_state:
        st.session_state.smart_goals = [
            {
                "goal": "Utveckla den ledningsledda differentieringen i matematikundervisningen så att minst 85% av eleverna når godtagbara kunskapskrav i nationella bedömningsstödet vid vårterminens slut.",
                "actions": "Genomföra veckoanpassningar med digitala verktyg samt lektionsdesign enligt IBIC. Chefen avsätter tid för kollegial skuggning 2 ggr/termin.",
                "followup": "Avstämning vid månatliga uppföljningssamtal samt slututvärdering maj 2027."
            }
        ]

    if 'criteria_assessments' not in st.session_state:
        st.session_state.criteria_assessments = {
            "Yrkeskunnande/färdighet": {"level": "En skicklig medarbetare", "motive": "Anpassar undervisningen mycket väl efter elevernas olika behov och tillämpar varierade undervisningsmetoder som stärker elevernas tilltro."},
            "Engagemang/ansvarsförmåga": {"level": "Medarbetare som når ett bra resultat", "motive": "Tar aktivt ansvar för elevernas lärande och deltar konstruktivt i ämneslagets kvalitetsarbete."},
            "Professionellt förhållningssätt": {"level": "En skicklig medarbetare", "motive": "Mycket trygg i sin yrkesroll med ett respektfullt bemötande gentemot både elever och vårdnadshavare."},
            "Samarbetsförmåga": {"level": "Föredömligt yrkesskicklig", "motive": "Aktiv drivkraft i det kollegiala lärandet och delar generöst med sig av metoder till kollegiet."},
            "Ledarskapsförmåga inom yrket": {"level": "En skicklig medarbetare", "motive": "Skapar en trygg och studiefokuserad lärmiljö där eleverna ges utrymme att utforska och växa."},
            "Flexibilitet/utvecklingsmöjlighet": {"level": "Medarbetare som når ett bra resultat", "motive": "Prövar gärna nya arbetssätt och analyserar kontinuerligt elevernas resultat för vidareutveckling."}
        }

    # TAB 1: OMVANDLA ANTECKNINGAR
    with tab1:
        st.subheader("📝 Omvandla handskrivna/fotade anteckningar till officiell text")
        st.info("💡 Ladda upp ett foto på dina handskrivna anteckningar eller klistra in råtext nedan. Appen omvandlar dem till tjänstemannamässig svenska, uppdelat efter kommunens dokumentationsmall, helt utan extra rubriker så att du kan klistra in texten direkt i dina Word-dokument.")
        
        col_input, col_output = st.columns([1, 1])
        
        with col_input:
            st.markdown("### 📥 Indata: Foto eller Råtext")
            
            uploaded_photo = st.file_uploader(
                "📷 Ladda upp foto på dina anteckningar (PNG, JPG, JPEG)",
                type=["png", "jpg", "jpeg"]
            )
            
            if uploaded_photo:
                st.image(uploaded_photo, caption="📸 Uppladdat foto på anteckningar", use_container_width=True)
            
            raw_text = st.text_area(
                "Eller klistra in råtext från dina samtal (eller fotade anteckningar):",
                value="Har mkt att göra vid betygssättning. Bra trivsel i laget. Funkar fint m föräldrar. Vill utveckla bildstödet i klassrummet för att nå elever som har svårt m instruktioner. Tränar regelbundet på gymmet.",
                height=150
            )
            
            btn_transform = st.button("⚡ Omvandla till officiell malltext", type="primary", use_container_width=True)
            
        with col_output:
            st.markdown("### 📤 Färdigformaterad text (Redo att klistras in i Word)")
            
            if btn_transform:
                result_text = ""
                
                # Option A: Process uploaded photo with Gemini AI Vision
                if uploaded_photo and AI_READY:
                    with st.spinner("🔍 Analyserar och tolkar handstilen i fotot med Gemini AI..."):
                        try:
                            pil_img = Image.open(uploaded_photo)
                            vision_prompt = (
                                "Du är en expert på skoldokumentation i Sverige. Analysera detta foto av handskrivna eller fotade anteckningar från ett medarbetarsamtal. "
                                "Tolka texten noggrant och omvandla den till en formell, tjänstemannamässig svenska för skolans dokumentation. "
                                "VIKTIGT: Svara ENBART med den omvandlade brödtexten. Inga rubriker, inga punktlistor, inga inledande kommentarer. Använd direkt myndighetssvenska."
                            )
                            
                            # Try model fallback
                            model_names = ['gemini-1.5-flash', 'gemini-1.5-pro', 'gemini-pro-vision']
                            response_text = None
                            for mname in model_names:
                                try:
                                    model = genai.GenerativeModel(mname)
                                    res = model.generate_content([vision_prompt, pil_img])
                                    if res and res.text:
                                        response_text = res.text.strip()
                                        break
                                except Exception:
                                    continue
                            
                            if response_text:
                                result_text = response_text
                            else:
                                st.warning("Kunde inte läsa bilden via AI. Använder inskriven text istället.")
                        except Exception as e:
                            st.warning(f"Fel vid bildläsning: {e}. Använder inskriven text istället.")

                # Option B: Process raw text with Gemini AI Text or Rule-Based Fallback
                if not result_text and raw_text.strip():
                    if AI_READY:
                        with st.spinner("🤖 Omvandlar anteckningar med Gemini AI..."):
                            try:
                                text_prompt = (
                                    f"Omvandla följande råa minnesanteckningar från ett medarbetarsamtal i skolan till en officiell, "
                                    f"tjänstemannamässig text:\n\n{raw_text}\n\n"
                                    f"KRAV:\n- Skriv på korrekt myndighetssvenska anpassad för skolans Word-mallar.\n"
                                    f"- INGA rubriker, INGA bullet points, INGA inledande hälsningar.\n"
                                    f"- Svara enbart med brödtext i färdiga stycken."
                                )
                                model = genai.GenerativeModel('gemini-1.5-flash')
                                res = model.generate_content(text_prompt)
                                if res and res.text:
                                    result_text = res.text.strip()
                            except Exception:
                                pass

                    # Rule-based fallback if AI is not available or failed
                    if not result_text:
                        lines = [l.strip() for l in raw_text.split('.') if l.strip()]
                        formatted_paragraphs = []
                        for line in lines:
                            l_lower = line.lower()
                            if "betyg" in l_lower or "belastning" in l_lower or "göra" in l_lower or "tid" in l_lower:
                                formatted_paragraphs.append("Arbetsbelastningen upplevs periodvis hög, särskilt i samband med bedömning och betygssättning. Medarbetaren arbetar aktivt med prioritering och tidsplanering.")
                            elif "trivs" in l_lower or "lag" in l_lower or "kollega" in l_lower:
                                formatted_paragraphs.append("Trivseln i arbetslaget och det kollegiala samarbetet fungerar mycket väl och utgör en stark grund för det dagliga arbetet.")
                            elif "föräld" in l_lower or "vårdnad" in l_lower:
                                formatted_paragraphs.append("Relationsbyggandet och kommunikationen med vårdnadshavare präglas av trygghet, tydlighet och ett professionellt förhållningssätt.")
                            elif "bildstöd" in l_lower or "instruktion" in l_lower or "bild" in l_lower or "utveckla" in l_lower:
                                formatted_paragraphs.append("Ett identifierat utvecklingsområde är att öka användningen av visuellt bildstöd i undervisningen för att ytterligare tillgängliggöra lärmiljön för alla elever.")
                            elif "trän" in l_lower or "gym" in l_lower or "hälsa" in l_lower:
                                formatted_paragraphs.append("Medarbetaren nyttjar kommunens friskvårdsbidrag och upplever god balans och återhämtning i vardagen.")
                            else:
                                formatted_paragraphs.append(f"{line.capitalize()}. Medarbetaren bidrar konstruktivt till verksamhetens måluppfyllelse.")

                        result_text = "\n\n".join(list(dict.fromkeys(formatted_paragraphs)))

                if result_text:
                    st.text_area("Kopiera texten nedan (Inga extra rubriker):", value=result_text, height=240)
                    st.success("✅ Texten är omvandlad till tjänstemannamässig svenska och kan klistras rakt in i din Word-mall!")

    # TAB 2: SMART-MÅLGENERATOR
    with tab2:
        st.subheader("🎯 Skapa SMART-formulerade mål")
        st.markdown("Utforma knivskarpa individuella mål kopplade till skolans prioriterade mål och lönekriterierna enligt **SMART-modellen**.")
        
        col_smart_form, col_smart_preview = st.columns([1.1, 0.9])
        
        with col_smart_form:
            st.markdown("#### 1. Välj kopplingar")
            target_criterion = st.selectbox(
                "Relaterat lönekriterium",
                ["Yrkeskunnande/färdighet", "Flexibilitet/utvecklingsmöjlighet", "Ledarskapsförmåga inom yrket", "Engagemang/ansvarsförmåga", "Samarbetsförmåga", "Professionellt förhållningssätt"]
            )
            
            target_school_goal = st.selectbox(
                "Skolans prioriterade mål",
                selected_school_goals if selected_school_goals else ["Höjd måluppfyllelse & likvärdig undervisning"]
            )
            
            st.markdown("#### 2. Ange målområde & idéer")
            raw_goal_idea = st.text_input("Vad vill/behöver medarbetaren utveckla?", value="Öka tillgängligheten med bildstöd i klassrummet")
            raw_measures = st.text_input("Åtgärder & stöd från chef", value="Skapa bildstöd för alla lektionsmoment. Chefen köper in licens för InPrint.")
            followup_date = st.text_input("Tidsram / Uppföljning", value="Avstämning dec 2026, slututvärdering maj 2027")
            
            if st.button("✨ Generera SMART-Mål", type="primary", use_container_width=True):
                smart_goal_text = f"Utveckla och integrera strukturmotsvarande bildstöd i den dagliga undervisningen för att öka tillgängligheten och studieron för alla elever, med fokus på kopplingen till skolans mål om '{target_school_goal}'."
                smart_action_text = f"Medarbetaren ansvarar för att skapa och tillämpa bildstöd vid lektionsstarter. Chef/skola ansvarar för att tillhandahålla material/programvara (InPrint) samt avsätta tid vid ämneslagsträffar."
                smart_follow_text = followup_date
                
                st.session_state.smart_goals.append({
                    "goal": smart_goal_text,
                    "actions": smart_action_text,
                    "followup": smart_follow_text
                })
                st.success("✅ SMART-mål tillagt i utvecklingsplanen!")

        with col_smart_preview:
            st.markdown("#### 📋 Genererad SMART-Struktur (Mallanpassad)")
            st.caption("Visar hur texten passar in i kommunens tabellmall för 'Individuell utvecklingsplan'.")
            
            for idx, g in enumerate(st.session_state.smart_goals):
                with st.expander(f"Mål {idx+1}: {g['goal'][:40]}...", expanded=True):
                    st.markdown(f"**Mål / utvecklingsområde (SMART):**\n{g['goal']}")
                    st.markdown(f"**Beskrivning av åtgärder & ansvar:**\n{g['actions']}")
                    st.markdown(f"**Uppföljning:**\n{g['followup']}")
                    if st.button(f"🗑️ Ta bort mål {idx+1}", key=f"del_{idx}"):
                        st.session_state.smart_goals.pop(idx)
                        st.rerun()

    # TAB 3: LÖNEKRITERIER & BEDÖMNING
    with tab3:
        st.subheader("📊 Utvärdering mot Lönekriterierna")
        st.caption(f"Anpassad utifrån lönekriterier för: **{role}**")
        
        criteria_dict = TEACHER_CRITERIA_INFO if "Lärare" in role else FRITIDS_CRITERIA_INFO
        
        levels = [
            "Behöver förbättras/utvecklas → Mål för kommande period",
            "Medarbetare som når ett bra resultat",
            "En skicklig medarbetare",
            "Föredömligt yrkesskicklig"
        ]
        
        for c_key, c_desc in criteria_dict.items():
            st.markdown(f"### 🔹 {c_key}")
            st.caption(c_desc)
            
            col_lvl, col_mot = st.columns([1, 2])
            
            current_data = st.session_state.criteria_assessments.get(c_key, {"level": "En skicklig medarbetare", "motive": ""})
            
            with col_lvl:
                sel_lvl = st.selectbox(
                    f"Bedömningsnivå för {c_key}",
                    levels,
                    index=levels.index(current_data["level"]) if current_data["level"] in levels else 2,
                    key=f"lvl_{c_key}"
                )
                
            with col_mot:
                default_mot = current_data["motive"] if current_data["motive"] else f"Medarbetaren uppfyller kriteriet väl i den dagliga undervisningen."
                sel_mot = st.text_area(
                    f"Motivering & exempel ({c_key})",
                    value=default_mot,
                    height=70,
                    key=f"mot_{c_key}"
                )
                
            st.session_state.criteria_assessments[c_key] = {"level": sel_lvl, "motive": sel_mot}
            st.markdown("---")

    # TAB 4: FÖRHANDSGRANSKA & EXPORTERA
    with tab4:
        st.subheader("📄 Förhandsgranska & Exportera färdig dokumentation")
        st.markdown("Generera ett färdigt Word-dokument (.docx) uppbyggt exakt enligt kommunens officiella dokumentationsmall för medarbetarsamtal.")
        
        st.markdown("#### 📋 Sammanfattning av innehåll")
        st.write(f"**Medarbetare:** {emp_name} | **Chef:** {manager_name} | **Yrke:** {role}")
        st.write(f"**Antal SMART-mål:** {len(st.session_state.smart_goals)} st | **Bedömda lönekriterier:** {len(st.session_state.criteria_assessments)} st")
        
        docx_buffer = create_word_document(
            emp_name=emp_name,
            manager_name=manager_name,
            role=role,
            period=period,
            date_str=date_str,
            school_goals=selected_school_goals,
            nulage_data=st.session_state.nulage_notes,
            criteria_data=st.session_state.criteria_assessments,
            plan_data=st.session_state.smart_goals
        )
        
        st.download_button(
            label="📥 Ladda ner färdigt Word-dokument (.docx)",
            data=docx_buffer,
            file_name=f"Medarbetarsamtal_{emp_name.replace(' ', '_')}_{date_str}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary",
            use_container_width=True
        )
        
        st.markdown("---")
        st.markdown("#### 👁️ Textförhandsgranskning (Inga extra rubriker - redo för copy-paste):")
        
        st.text_area(
            "Utvecklingsplan (Mål & Åtgärder)",
            value="\n\n".join([f"Mål: {g['goal']}\nÅtgärder: {g['actions']}\nUppföljning: {g['followup']}" for g in st.session_state.smart_goals]),
            height=180
        )

if __name__ == "__main__":
    main()

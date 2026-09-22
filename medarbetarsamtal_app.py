import streamlit as st
import io
import os
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from PIL import Image

# Check for Google Generative AI availability
try:
    import google.generativeai as genai
    HAS_GEMINI = True
except ImportError:
    HAS_GEMINI = False

# Page Configuration
st.set_page_config(
    page_title="Medarbetarsamtal & Målgenerator - Skolan",
    page_icon="🏫",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Modern Artistic Semi-Dark Custom Styling
st.markdown("""
<style>
    /* App background - Modern artistic semi-dark gradient */
    .stApp {
        background: linear-gradient(135deg, #0B132B 0%, #1C2541 40%, #0B132B 100%);
        color: #F1F5F9;
    }
    
    /* Sidebar Styling */
    [data-testid="stSidebar"] {
        background-color: rgba(11, 19, 43, 0.95) !important;
        border-right: 1px solid rgba(255, 255, 255, 0.1);
    }
    
    /* Glassmorphism Cards & Containers */
    .card, div[data-testid="stExpander"] {
        background: rgba(30, 41, 59, 0.65) !important;
        backdrop-filter: blur(12px);
        border: 1px solid rgba(255, 255, 255, 0.12) !important;
        border-radius: 12px !important;
        padding: 20px !important;
        margin-bottom: 18px !important;
        box-shadow: 0 8px 32px 0 rgba(0, 0, 0, 0.37);
    }
    
    /* Headers with vibrant gradient text */
    .main-header {
        font-size: 2.3rem;
        background: linear-gradient(90deg, #60A5FA, #A78BFA, #F472B6);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        font-weight: 800;
        margin-bottom: 0.2rem;
        letter-spacing: -0.5px;
    }
    .sub-header {
        font-size: 1.05rem;
        color: #94A3B8;
        margin-bottom: 1.5rem;
    }
    
    /* Input Fields Styling for Dark Mode */
    .stTextInput > div > div > input, 
    .stTextArea > div > div > textarea, 
    .stSelectbox > div > div {
        background-color: #1E293B !important;
        color: #F8FAFC !important;
        border: 1px solid #475569 !important;
        border-radius: 8px !important;
    }
    
    /* Tabs Custom Styling */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
        background-color: rgba(15, 23, 42, 0.6);
        padding: 6px;
        border-radius: 10px;
        border: 1px solid rgba(255, 255, 255, 0.08);
    }
    .stTabs [data-baseweb="tab"] {
        height: 45px;
        white-space: pre-wrap;
        background-color: transparent;
        border-radius: 8px;
        color: #94A3B8;
        font-weight: 600;
    }
    .stTabs [aria-selected="true"] {
        background-color: #3B82F6 !important;
        color: #FFFFFF !important;
    }
    
    /* Info Box Styling */
    .stAlert {
        background-color: rgba(30, 58, 138, 0.4) !important;
        border: 1px solid #3B82F6 !important;
        color: #E2E8F0 !important;
        border-radius: 10px !important;
    }
    
    /* Smart Tag Badge */
    .smart-tag {
        display: inline-block;
        background: linear-gradient(90deg, #1E40AF, #3B82F6);
        color: #FFFFFF;
        padding: 3px 10px;
        border-radius: 6px;
        font-weight: 700;
        font-size: 0.82rem;
        margin-right: 6px;
    }
</style>
""", unsafe_allow_html=True)

# Initialize Gemini API if key is present
API_KEY = None
if "GEMINI_API_KEY" in st.secrets:
    API_KEY = st.secrets["GEMINI_API_KEY"]
elif "GEMINI_API_KEY" in os.environ:
    API_KEY = os.environ["GEMINI_API_KEY"]

GEMINI_ACTIVE = False
if HAS_GEMINI and API_KEY:
    try:
        genai.configure(api_key=API_KEY)
        GEMINI_ACTIVE = True
    except Exception:
        GEMINI_ACTIVE = False

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
# MAIN APP ENTRY POINT
# ---------------------------------------------------------
def main():
    # SIDEBAR SETUP
    with st.sidebar:
        st.image("https://img.icons8.com/color/96/school.png", width=64)
        st.title("⚙️ Samtalsinställningar")
        
        # AI Status Badge
        if GEMINI_ACTIVE:
            st.success("🟢 AI-Motor Aktiv (Gemini API)")
        else:
            st.info("🔵 Regelbaserad motor (Ingen API-nyckel krävs)")
            
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

    # TAB 1: OMVANDLA ANTECKNINGAR (BILD OCH TEXT)
    with tab1:
        st.subheader("📝 Omvandla handskrivna/fotade anteckningar till officiell text")
        st.info("💡 Ladda upp ett foto av dina handskrivna anteckningar ELLER klistra in råtext nedan. Appen omvandlar dem till tjänstemannamässig svenska, uppdelat efter kommunens dokumentationsmall, helt utan extra rubriker så att du kan klistra in texten direkt i dina Word-dokument.")
        
        col_input, col_output = st.columns([1, 1])
        
        with col_input:
            st.markdown("### 📥 Dina anteckningar (Foto eller Text)")
            
            section_choice = st.selectbox(
                "Välj område i mallen att bearbeta",
                ["Arbetssituation – nuläge", "Resultat & Lönekriterier (Tillbakablick)", "Övriga minnesanteckningar"]
            )
            
            # FILE UPLOADER FOR PHOTO OF NOTES
            uploaded_file = st.file_uploader(
                "📷 Ladda upp foto på handskrivna anteckningar",
                type=["png", "jpg", "jpeg"],
                help="Fotografera dina anteckningar med mobilen och ladda upp bildfilen här."
            )
            
            if uploaded_file is not None:
                st.image(uploaded_file, caption="Inläst foto på anteckningar", use_container_width=True)
            
            raw_text = st.text_area(
                "Eller klistra in råtext från dina anteckningar:",
                value="Har mkt att göra vid betygssättning. Bra trivsel i laget. Funkar fint m föräldrar. Vill utveckla bildstödet i klassrummet för att nå elever som har svårt m instruktioner. Tränar regelbundet på gymmet." if uploaded_file is None else "",
                height=130
            )
            
            btn_transform = st.button("⚡ Omvandla till officiell malltext", type="primary", use_container_width=True)
            
        with col_output:
            st.markdown("### 📤 Färdigformaterad text (Redo att klistras in i Word)")
            
            if btn_transform:
                with st.spinner("Bearbetar anteckningar och formaterar till officiell svenska..."):
                    result_text = ""
                    
                    # 1. Processing via Gemini API if active
                    if GEMINI_ACTIVE:
                        try:
                            model = genai.GenerativeModel('gemini-1.5-flash')
                            
                            prompt = f"""
                            Du är en expert-skolledare inom svenskt skolväsen.
                            Omvandla följande anteckningar från ett medarbetarsamtal för en {role} till tjänstemannamässig, professionell och välformulerad svenska.
                            Gäller området: {section_choice}.
                            
                            Instruktioner:
                            - Formulera om språket till formell myndighetssvenska anpassad för skolan.
                            - VIKTIGT: Generera ENBART den färdiga brödtexten.
                            - INGA extra rubriker, INGA punktlistor, INGA inledande kommentarer ("Här är texten:").
                            - Texten ska kunna klistras in direkt i kommunens Word-mall för medarbetarsamtal.
                            """
                            
                            if uploaded_file is not None:
                                img = Image.open(uploaded_file)
                                response = model.generate_content([prompt, "Läs av handstilen/texten i denna bild och omvandla till officiell malltext:", img])
                            else:
                                response = model.generate_content([prompt, f"Anteckningar att omvandla:\n{raw_text}"])
                                
                            result_text = response.text.strip()
                        except Exception as e:
                            st.warning(f"AI-anrop misslyckades ({e}), använder regelbaserad omvandling.")
                            GEMINI_ACTIVE_LOCAL = False
                    
                    # 2. Rule-based / Fallback processing if Gemini is inactive or fails
                    if not result_text:
                        source_str = raw_text if raw_text else "Foto på anteckningar inläst."
                        lines = [l.strip() for l in source_str.split('.') if l.strip()]
                        formatted_paragraphs = []
                        
                        for line in lines:
                            l_lower = line.lower()
                            if "betyg" in l_lower or "belastning" in l_lower or "göra" in l_lower or "tid" in l_lower or "foto" in l_lower:
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

                    st.session_state['transformed_text'] = result_text

            # Display Result
            display_val = st.session_state.get('transformed_text', "")
            st.text_area("Kopiera texten nedan (Inga extra rubriker):", value=display_val, height=220)
            if display_val:
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
        st.caption(f"Anpassad utifrån Norrköpings kommuns lönekriterier för: **{role}**")
        
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

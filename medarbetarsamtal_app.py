import streamlit as st
import io
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from PIL import Image

# Safe import for Google Generative AI
try:
    import google.generativeai as genai
    HAS_GENAI = True
except ImportError:
    HAS_GENAI = False

# Page Configuration
st.set_page_config(
    page_title="Medarbetarsamtal & Målgenerator - Skolan",
    page_icon="🏫",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom Styling - Modern dark theme with glowing geometric lines & maximum text contrast
st.markdown("""
<style>
    /* Dark background with modern geometric glowing line gradient */
    .stApp {
        background: 
            radial-gradient(circle at 15% 15%, rgba(99, 102, 241, 0.15) 0%, transparent 40%),
            radial-gradient(circle at 85% 85%, rgba(168, 85, 247, 0.15) 0%, transparent 40%),
            linear-gradient(135deg, #0B0F19 0%, #111827 50%, #0F172A 100%);
        background-attachment: fixed;
        color: #F8FAFC !important;
    }

    /* All typography contrast fixes */
    h1, h2, h3, h4, h5, h6, p, label, span, div {
        color: #F8FAFC !important;
    }
    
    .stCaption, .stCaption p {
        color: #94A3B8 !important;
    }

    .main-header {
        font-size: 2.2rem;
        background: linear-gradient(90deg, #60A5FA 0%, #A78BFA 50%, #F472B6 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        font-weight: 800;
        margin-bottom: 0.2rem;
    }

    .sub-header {
        font-size: 1.1rem;
        color: #CBD5E1 !important;
        margin-bottom: 1.5rem;
    }

    /* Frosted glass container style */
    .glass-card {
        background: rgba(30, 41, 59, 0.7);
        backdrop-filter: blur(12px);
        -webkit-backdrop-filter: blur(12px);
        border: 1px solid rgba(255, 255, 255, 0.1);
        border-radius: 12px;
        padding: 20px;
        margin-bottom: 18px;
        box-shadow: 0 8px 32px 0 rgba(0, 0, 0, 0.37);
    }

    /* Input elements styling */
    .stTextInput input, .stTextArea textarea, .stSelectbox select {
        background-color: #0F172A !important;
        color: #FFFFFF !important;
        border: 1px solid #334155 !important;
        border-radius: 8px !important;
    }

    .stTextInput input:focus, .stTextArea textarea:focus {
        border-color: #818CF8 !important;
        box-shadow: 0 0 0 2px rgba(129, 140, 248, 0.2) !important;
    }

    /* Buttons */
    .stButton>button {
        background: linear-gradient(90deg, #4F46E5 0%, #7C3AED 100%) !important;
        color: #FFFFFF !important;
        font-weight: 600 !important;
        border: none !important;
        border-radius: 8px !important;
        padding: 0.5rem 1rem !important;
        transition: all 0.2s ease-in-out !important;
    }
    
    .stButton>button:hover {
        transform: translateY(-1px);
        box-shadow: 0 4px 12px rgba(124, 58, 237, 0.4) !important;
    }

    /* Status badges */
    .status-badge-ok {
        background-color: rgba(16, 185, 129, 0.2);
        color: #34D399 !important;
        border: 1px solid rgba(52, 211, 153, 0.3);
        padding: 4px 10px;
        border-radius: 6px;
        font-weight: 600;
        font-size: 0.85rem;
    }

    .status-badge-info {
        background-color: rgba(59, 130, 246, 0.2);
        color: #60A5FA !important;
        border: 1px solid rgba(96, 165, 250, 0.3);
        padding: 4px 10px;
        border-radius: 6px;
        font-weight: 600;
        font-size: 0.85rem;
    }

    /* Sidebar dark styling */
    section[data-testid="stSidebar"] {
        background-color: #0F172A !important;
        border-right: 1px solid #1E293B !important;
    }

    /* Tabs styling */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
        background-color: rgba(15, 23, 42, 0.6);
        padding: 6px;
        border-radius: 10px;
    }

    .stTabs [data-baseweb="tab"] {
        color: #94A3B8 !important;
        border-radius: 6px;
        padding: 8px 16px;
        font-weight: 500;
    }

    .stTabs [aria-selected="true"] {
        background-color: #312E81 !important;
        color: #FFFFFF !important;
        font-weight: 700;
    }
</style>
""", unsafe_allow_html=True)

# ---------------------------------------------------------
# GEMINI API INITIALIZATION
# ---------------------------------------------------------
GEMINI_CONFIGURED = False
GEMINI_API_KEY = None

if HAS_GENAI:
    if "GEMINI_API_KEY" in st.secrets:
        GEMINI_API_KEY = st.secrets["GEMINI_API_KEY"]
    elif "google_api_key" in st.secrets:
        GEMINI_API_KEY = st.secrets["google_api_key"]
        
    if GEMINI_API_KEY:
        try:
            genai.configure(api_key=GEMINI_API_KEY)
            GEMINI_CONFIGURED = True
        except Exception:
            GEMINI_CONFIGURED = False

# ---------------------------------------------------------
# AI TRANSFORMATION FUNCTIONS (TEXT & VISION)
# ---------------------------------------------------------
def transform_photo_notes_with_ai(image, area_choice):
    """Processes a photo of handwritten notes directly using Gemini Vision and formats into official Swedish text."""
    if HAS_GENAI and GEMINI_CONFIGURED:
        try:
            model = genai.GenerativeModel('gemini-1.5-flash')
            prompt = f"""Du är en erfaren rektor och chef i svensk skola.
Läs av och tolka de handskrivna eller tryckta anteckningarna på denna bild från ett medarbetarsamtal.
Omvandla direkt bildens innehåll till tjänstemannamässig, formell och professionell svenska lämplig för kommunens skoldokumentation.

Område i samtalet: {area_choice}

Kritiska instruktioner:
1. Skriv på formell och tjänstemannamässig svenska anpassad för skolan.
2. Inga rubriker, inga inledande hälsningar ("Här är..."), inga punktlistor och inga citationstecken.
3. Texten ska vara färdigformaterad för att klistras in direkt i en Word-mall.
4. Behåll all fakta och alla punkter från anteckningarna i bilden men formulera dem professionellt."""
            
            response = model.generate_content([prompt, image])
            if response.text:
                return response.text.strip()
        except Exception as e:
            st.warning(f"⚠️ AI-bildtolkning misslyckades ({e}). Kontrollera din Gemini API-nyckel.")
            
    return "Kunde inte tolka bilden automatiskt. Se till att GEMINI_API_KEY finns i Streamlit Secrets för automatisk fototolkning."

def transform_notes_with_ai(raw_text, area_choice):
    """Transforms raw text notes into formal Swedish text for official school documentation."""
    if HAS_GENAI and GEMINI_CONFIGURED and raw_text.strip():
        try:
            model = genai.GenerativeModel('gemini-1.5-flash')
            prompt = f"""Du är en erfaren rektor och chef i svensk skola.
Omvandla följande råa/handskrivna anteckningar från ett medarbetarsamtal till tjänstemannamässig, formell och professionell svenska lämplig för kommunens skoldokumentation.

Område i samtalet: {area_choice}
Råa anteckningar:
{raw_text}

Kritiska instruktioner:
1. Skriv på formell och tjänstemannamässig svenska anpassad för skolan.
2. Inga rubriker, inga inledande hälsningar ("Här är..."), inga punktlistor och inga citationstecken.
3. Texten ska vara färdigformaterad för att klistras in direkt i en Word-mall.
4. Behåll fakta från anteckningarna men formulera dem professionellt."""
            
            response = model.generate_content(prompt)
            if response.text:
                return response.text.strip()
        except Exception as e:
            st.warning(f"⚠️ AI-anrop misslyckades ({e}). Använder inbyggd regelbaserad omvandling.")

    # Rule-based fallback
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

    return "\n\n".join(list(dict.fromkeys(formatted_paragraphs)))

def generate_smart_goal_with_ai(criterion, school_goal, raw_idea, raw_measures, followup):
    """Generates a structured SMART goal using AI or fallback rules."""
    if HAS_GENAI and GEMINI_CONFIGURED and raw_idea.strip():
        try:
            model = genai.GenerativeModel('gemini-1.5-flash')
            prompt = f"""Du är rektor och expert på SMART-modellen i skolan.
Skapa ett konkret SMART-mål (Specifikt, Mätbart, Accepterat, Realistiskt, Tidsatt) baserat på följande uppgifter:

Lönekriterium: {criterion}
Skolans prioriterade mål: {school_goal}
Utvecklingsområde/Idé: {raw_idea}
Åtgärder & Stöd: {raw_measures}
Tidsram: {followup}

Svara exakt i följande tre rader:
MÅL: [Det SMART-formulerade målet]
ÅTGÄRDER: [Åtgärder samt tydlig ansvarsfördelning mellan medarbetare och chef/skola]
UPPFÖLJNING: [Uppföljningsplan och tidsram]"""
            
            response = model.generate_content(prompt)
            if response.text:
                lines = response.text.strip().split('\n')
                g_text, a_text, f_text = "", "", followup
                for l in lines:
                    if l.startswith("MÅL:"):
                        g_text = l.replace("MÅL:", "").strip()
                    elif l.startswith("ÅTGÄRDER:"):
                        a_text = l.replace("ÅTGÄRDER:", "").strip()
                    elif l.startswith("UPPFÖLJNING:"):
                        f_text = l.replace("UPPFÖLJNING:", "").strip()
                if g_text and a_text:
                    return g_text, a_text, f_text
        except Exception:
            pass

    # Fallback rules
    smart_goal_text = f"Utveckla och integrera {raw_idea.lower()} i den dagliga undervisningen för att öka tillgängligheten och studieron för alla elever, kopplat till skolans prioriterade mål om '{school_goal}'."
    smart_action_text = f"Medarbetaren ansvarar för att planera och genomföra åtgärderna ({raw_measures}). Chef/skola ansvarar för resurser, kompetensutveckling samt förutsättningar vid ämneslagsträffar."
    return smart_goal_text, smart_action_text, followup

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
    r_title.font.color.rgb = RGBColor(0x1E, 0x3A, 0x8A)
    
    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_sub = p_sub.add_run(f"Norrköpings kommun - Skolverksamhet | {period}")
    r_sub.font.size = Pt(11)
    r_sub.font.italic = True
    
    doc.add_paragraph() # Spacer

    # Info Table
    info_table = doc.add_table(rows=3, cols=2)
    info_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    info_table.autofit = False
    
    col_widths = [Inches(3.2), Inches(3.2)]
    
    info_data = [
        [f"Medarbetare: {emp_name}", f"Chef/Samtalsledare: {manager_name}"],
        [f"Yrkestitel: {role}", f"Samtalsdatum: {date_str}"],
        [f"Tidsperiod: {period}", f"Skolans mål: {', '.join(school_goals[:2])}"]
    ]
    
    for row_idx, row in enumerate(info_table.rows):
        for col_idx, cell in enumerate(row.cells):
            cell.width = col_widths[col_idx]
            cell.text = info_data[row_idx][col_idx]
            set_cell_background(cell, "F1F5F9")
            
    doc.add_paragraph() # Spacer

    # Section 1: Nuläge & Arbetssituation
    h1 = doc.add_heading("1. Arbetssituation och Nuläge", level=1)
    h1.style.font.color.rgb = RGBColor(0x1E, 0x3A, 0x8A)
    
    for topic, text in nulage_data.items():
        p_topic = doc.add_paragraph()
        r_t = p_topic.add_run(f"• {topic}: ")
        r_t.bold = True
        p_topic.add_run(text)

    doc.add_paragraph()

    # Section 2: Individuell Utvecklingsplan (SMART-Mål)
    h2 = doc.add_heading("2. Individuell Utvecklingsplan (Mål & Åtgärder)", level=1)
    h2.style.font.color.rgb = RGBColor(0x1E, 0x3A, 0x8A)
    
    if plan_data:
        plan_table = doc.add_table(rows=1, cols=3)
        plan_table.autofit = False
        
        # Headers
        hdr_cells = plan_table.rows[0].cells
        headers = ["Mål / utvecklingsområde (SMART)", "Åtgärder & Ansvarsfördelning", "Uppföljningsplan"]
        widths = [Inches(2.5), Inches(2.5), Inches(1.5)]
        
        for idx, text in enumerate(headers):
            hdr_cells[idx].width = widths[idx]
            hdr_cells[idx].text = text
            set_cell_background(hdr_cells[idx], "1E3A8A")
            for p in hdr_cells[idx].paragraphs:
                for r in p.runs:
                    r.font.bold = True
                    r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
                    
        for goal_item in plan_data:
            row_cells = plan_table.add_row().cells
            row_cells[0].width = widths[0]
            row_cells[0].text = goal_item["goal"]
            
            row_cells[1].width = widths[1]
            row_cells[1].text = goal_item["actions"]
            
            row_cells[2].width = widths[2]
            row_cells[2].text = goal_item["followup"]
            
            for cell in row_cells:
                set_cell_background(cell, "FAFAFA")
    else:
        doc.add_paragraph("Inga specifika mål registrerade för perioden.")

    doc.add_paragraph()

    # Section 3: Utvärdering av Lönekriterier
    h3 = doc.add_heading("3. Utvärdering mot Lönekriterier", level=1)
    h3.style.font.color.rgb = RGBColor(0x1E, 0x3A, 0x8A)
    
    crit_table = doc.add_table(rows=1, cols=3)
    crit_table.autofit = False
    
    c_hdr_cells = crit_table.rows[0].cells
    c_headers = ["Lönekriterium", "Bedömd Nivå", "Motivering & Undervisningsexempel"]
    c_widths = [Inches(1.8), Inches(1.8), Inches(2.9)]
    
    for idx, text in enumerate(c_headers):
        c_hdr_cells[idx].width = c_widths[idx]
        c_hdr_cells[idx].text = text
        set_cell_background(c_hdr_cells[idx], "1E3A8A")
        for p in c_hdr_cells[idx].paragraphs:
            for r in p.runs:
                r.font.bold = True
                r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
                
    for c_name, c_val in criteria_data.items():
        row_cells = crit_table.add_row().cells
        row_cells[0].width = c_widths[0]
        row_cells[0].text = c_name
        
        row_cells[1].width = c_widths[1]
        row_cells[1].text = c_val["level"]
        
        row_cells[2].width = c_widths[2]
        row_cells[2].text = c_val["motive"]
        
        for cell in row_cells:
            set_cell_background(cell, "FFFFFF")

    doc.add_paragraph()
    
    # Signatures
    p_sig = doc.add_paragraph()
    p_sig.add_run("Underskrifter:\n\n\n___________________________               ___________________________\nMedarbetare                                Chef / Skolledare").bold = True

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# ---------------------------------------------------------
# CRITERIA DATA FROM MUNICIPALITY DOCUMENTS
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
# MAIN APPLICATION
# ---------------------------------------------------------
def main():
    # SIDEBAR CONFIGURATION
    with st.sidebar:
        st.image("https://img.icons8.com/color/96/school.png", width=70)
        st.title("⚙️ Inställningar")
        
        # Status Badge
        if GEMINI_CONFIGURED:
            st.markdown('<span class="status-badge-ok">🟢 AI (Gemini) Aktiverad</span>', unsafe_allow_html=True)
        else:
            st.markdown('<span class="status-badge-info">🔵 Regelbaserat läge</span>', unsafe_allow_html=True)
            if not HAS_GENAI:
                st.caption("ℹ️ Paket `google-generativeai` saknas i Streamlit Cloud.")
            elif not GEMINI_API_KEY:
                st.caption("ℹ️ Lägg till `GEMINI_API_KEY` i Streamlit Secrets för full AI-kraft.")

        st.markdown("---")
        st.subheader("👤 Samtalsdata")
        emp_name = st.text_input("Medarbetarens namn", value="Anna Andersson")
        manager_name = st.text_input("Chefens/Rektorns namn", value="Karin Skolledare")
        role = st.selectbox(
            "Yrkestitel",
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
    st.markdown('<div class="sub-header">Omvandla fotade eller råa anteckningar till officiell, färdigformaterad dokumentation och knivskarpa SMART-mål redo för kommunens Word-mallar.</div>', unsafe_allow_html=True)

    # App Navigation Tabs
    tab1, tab2, tab3, tab4 = st.tabs([
        "📝 Omvandla Anteckningar (Foto / Text)",
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
                "goal": "Utveckla och integrera strukturmotsvarande bildstöd i den dagliga undervisningen för att öka tillgängligheten och studieron för alla elever, kopplat till skolans prioriterade mål om 'Trygghet, studiero & tillgänglig lärmiljö'.",
                "actions": "Medarbetaren ansvarar för att skapa och tillämpa bildstöd vid lektionsstarter. Chef/skola ansvarar för att tillhandahålla material/programvara samt avsätta tid vid ämneslagsträffar.",
                "followup": "Avstämning december 2026, slututvärdering maj 2027."
            }
        ]

    if 'criteria_assessments' not in st.session_state:
        st.session_state.criteria_assessments = {
            "Yrkeskunnande/färdighet": {"level": "En skicklig medarbetare", "motive": "Anpassar undervisningen mycket väl efter elevernas olika behov och tillämpar varierade undervisningsmetoder som stärka elevernas tilltro."},
            "Engagemang/ansvarsförmåga": {"level": "Medarbetare som når ett bra resultat", "motive": "Tar aktivt ansvar för elevernas lärande och deltar konstruktivt i ämneslagets kvalitetsarbete."},
            "Professionellt förhållningssätt": {"level": "En skicklig medarbetare", "motive": "Mycket trygg i sin yrkesroll med ett respektfullt bemötande gentemot både elever och vårdnadshavare."},
            "Samarbetsförmåga": {"level": "Föredömligt yrkesskicklig", "motive": "Aktiv drivkraft i det kollegiala lärandet och delar generöst med sig av metoder till kollegiet."},
            "Ledarskapsförmåga inom yrket": {"level": "En skicklig medarbetare", "motive": "Skapar en trygg och studiefokuserad lärmiljö där eleverna ges utrymme att utforska och växa."},
            "Flexibilitet/utvecklingsmöjlighet": {"level": "Medarbetare som når ett bra resultat", "motive": "Prövar gärna nya arbetssätt och analyserar kontinuerligt elevernas resultat för vidareutveckling."}
        }

    # TAB 1: OMVANDLA ANTECKNINGAR (DIRECT PHOTO PROCESSING)
    with tab1:
        st.subheader("📝 Omvandla handskrivna eller fotade anteckningar till officiell text")
        st.info("💡 **Direkt fototolkning:** Ladda upp ett foto på dina handskrivna anteckningar nedan. Appen läser av bilden automatiskt och omvandlar innehållet direkt till tjänstemannamässig svenska – helt utan att du behöver skriva in någon text manuellt!")
        
        section_choice = st.selectbox(
            "Välj vilket område anteckningarna gäller",
            ["Arbetssituation – nuläge", "Resultat & Lönekriterier (Tillbakablick)", "Övriga minnesanteckningar"]
        )

        col_input, col_output = st.columns([1, 1])
        
        with col_input:
            st.markdown("### 📷 Ladda upp Foto på Anteckningar")
            uploaded_photo = st.file_uploader(
                "Välj bildfil (eller ta ett foto med mobilen):",
                type=["png", "jpg", "jpeg", "webp"],
                help="Appen tolkar handstil och tryckt text direkt från bilden."
            )
            
            if uploaded_photo is not None:
                try:
                    img = Image.open(uploaded_photo)
                    st.image(img, caption="Uppladdat foto på anteckningar", use_container_width=True)
                except Exception as e:
                    st.error(f"Kunde inte öppna bilden: {e}")

            st.markdown("---")
            with st.expander("✍️ Har du ingen bild? Klistra in text manuellt här (valfritt)"):
                manual_text = st.text_area(
                    "Manuell råtext (endast om du inte har ett foto):",
                    value="",
                    height=100
                )

            btn_transform = st.button("⚡ Omvandla till officiell malltext", type="primary", use_container_width=True)

        with col_output:
            st.markdown("### 📤 Färdigformaterad text (Redo för Word-mallen)")
            
            result_text = ""
            if btn_transform:
                if uploaded_photo is not None:
                    with st.spinner("🔍 Läser av foto och omvandlar handskriven text med AI..."):
                        img = Image.open(uploaded_photo)
                        result_text = transform_photo_notes_with_ai(img, section_choice)
                elif manual_text.strip():
                    with st.spinner("⏳ Omvandlar text..."):
                        result_text = transform_notes_with_ai(manual_text, section_choice)
                else:
                    st.warning("⚠️ Ladda upp ett foto eller skriv in text först!")

            if result_text:
                st.text_area("Kopiera texten nedan (Inga extra rubriker):", value=result_text, height=260)
                st.success("✅ Klart! Texten är omvandlad till tjänstemannamässig svenska och kan klistras rakt in i din Word-mall.")
            else:
                st.text_area("Kopiera texten nedan (Inga extra rubriker):", value="Här visas den omvandlade texten så fort du laddat upp ett foto och klickat på knappen ovan...", height=260)

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
                smart_goal_text, smart_action_text, smart_follow_text = generate_smart_goal_with_ai(
                    target_criterion, target_school_goal, raw_goal_idea, raw_measures, followup_date
                )
                
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

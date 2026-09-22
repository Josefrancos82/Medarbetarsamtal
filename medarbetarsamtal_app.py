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

# ---------------------------------------------------------
# COMPREHENSIVE LIGHT THEME STYLING
# ---------------------------------------------------------
st.markdown("""
<style>
    /* 1. Global App Viewport & Background */
    html, body, [data-testid="stAppViewContainer"], .stApp {
        background-color: #F8FAFC !important;
        background-image: 
            radial-gradient(at 0% 0%, rgba(219, 234, 254, 0.6) 0px, transparent 50%),
            radial-gradient(at 100% 100%, rgba(224, 231, 255, 0.6) 0px, transparent 50%),
            linear-gradient(135deg, #F1F5F9 0%, #FFFFFF 50%, #F8FAFC 100%) !important;
        background-attachment: fixed !important;
        color: #0F172A !important;
        font-family: 'Segoe UI', -apple-system, BlinkMacSystemFont, Roboto, sans-serif !important;
    }

    /* 2. Top Header Bar */
    [data-testid="stHeader"] {
        background-color: rgba(248, 250, 252, 0.85) !important;
        backdrop-filter: blur(8px) !important;
    }

    /* 3. Sidebar Styling */
    [data-testid="stSidebar"] {
        background-color: #FFFFFF !important;
        border-right: 1px solid #E2E8F0 !important;
        box-shadow: 2px 0 10px rgba(0,0,0,0.02) !important;
    }
    [data-testid="stSidebar"] * {
        color: #0F172A !important;
    }

    /* 4. Headings & Typography */
    .main-header {
        font-size: 2.2rem;
        color: #1E3A8A !important;
        font-weight: 800;
        margin-bottom: 0.2rem;
        letter-spacing: -0.02em;
    }
    .sub-header {
        font-size: 1.1rem;
        color: #334155 !important;
        margin-bottom: 1.5rem;
    }
    h1, h2, h3, h4, h5, h6, p, span, label, div {
        color: #0F172A !important;
    }
    .stMarkdown p, .stMarkdown label {
        color: #0F172A !important;
        font-weight: 500;
    }

    /* 5. Custom Cards */
    .card {
        background-color: #FFFFFF !important;
        border: 1px solid #E2E8F0 !important;
        border-radius: 12px !important;
        padding: 20px !important;
        margin-bottom: 15px !important;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05), 0 2px 4px -1px rgba(0, 0, 0, 0.03) !important;
    }

    /* 6. Inputs: Text Inputs, Text Areas, Selectboxes */
    .stTextInput input, .stTextArea textarea, div[data-baseweb="select"] > div {
        background-color: #FFFFFF !important;
        color: #000000 !important;
        border: 1px solid #CBD5E1 !important;
        border-radius: 8px !important;
        font-size: 0.95rem !important;
        font-weight: 500 !important;
    }
    .stTextInput input:focus, .stTextArea textarea:focus {
        border-color: #2563EB !important;
        box-shadow: 0 0 0 3px rgba(37, 99, 235, 0.15) !important;
    }

    /* Selectbox Dropdown Menu Popup */
    div[data-baseweb="popover"], div[data-baseweb="menu"], ul[data-baseweb="menu"] {
        background-color: #FFFFFF !important;
        color: #0F172A !important;
        border: 1px solid #E2E8F0 !important;
        border-radius: 8px !important;
    }
    li[data-baseweb="option"] {
        background-color: #FFFFFF !important;
        color: #0F172A !important;
    }
    li[data-baseweb="option"]:hover, li[data-baseweb="option"][aria-selected="true"] {
        background-color: #EFF6FF !important;
        color: #1D4ED8 !important;
    }

    /* 7. Tabs */
    .stTabs [data-baseweb="tab-list"] {
        background-color: #E2E8F0 !important;
        border-radius: 10px !important;
        padding: 4px !important;
        gap: 4px !important;
    }
    .stTabs [data-baseweb="tab"] {
        color: #475569 !important;
        font-weight: 600 !important;
        border-radius: 8px !important;
        padding: 8px 16px !important;
        background-color: transparent !important;
    }
    .stTabs [aria-selected="true"] {
        background-color: #FFFFFF !important;
        color: #1E3A8A !important;
        box-shadow: 0 2px 4px rgba(0,0,0,0.06) !important;
    }

    /* 8. Buttons */
    .stButton > button, .stDownloadButton > button {
        background-color: #1E3A8A !important;
        color: #FFFFFF !important;
        font-weight: 600 !important;
        border-radius: 8px !important;
        border: none !important;
        padding: 0.6rem 1.2rem !important;
        transition: all 0.2s ease !important;
    }
    .stButton > button:hover, .stDownloadButton > button:hover {
        background-color: #2563EB !important;
        color: #FFFFFF !important;
        box-shadow: 0 4px 12px rgba(37, 99, 235, 0.25) !important;
    }

    /* 9. Expanders */
    .streamlit-expanderHeader {
        background-color: #FFFFFF !important;
        color: #0F172A !important;
        border: 1px solid #E2E8F0 !important;
        border-radius: 8px !important;
        font-weight: 600 !important;
    }
    .streamlit-expanderContent {
        background-color: #F8FAFC !important;
        color: #0F172A !important;
        border: 1px solid #E2E8F0 !important;
        border-top: none !important;
    }

    /* 10. File Uploader */
    [data-testid="stFileUploader"] {
        background-color: #FFFFFF !important;
        border: 2px dashed #CBD5E1 !important;
        border-radius: 10px !important;
        padding: 10px !important;
    }

    /* 11. Custom Status Badges */
    .status-badge-ok {
        background-color: #DEF7EC !important;
        color: #03543F !important;
        padding: 4px 10px;
        border-radius: 6px;
        font-weight: 600;
        font-size: 0.85rem;
    }
    .status-badge-info {
        background-color: #E1EFFE !important;
        color: #1E40AF !important;
        padding: 4px 10px;
        border-radius: 6px;
        font-weight: 600;
        font-size: 0.85rem;
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
# CLEANING AI OUTPUT (REMOVES THINKING / METADATA)
# ---------------------------------------------------------
def clean_ai_output(text):
    """Filters out AI meta-comments, role descriptions, and process reasoning."""
    if not text:
        return ""
    lines = text.strip().split('\n')
    cleaned = []
    skip_prefixes = [
        'role:', 'task:', 'topic:', 'input:', 'constraints:', 
        'top part:', 'main section:', 're-evaluating', 'point 1:', 
        'correction:', 'refinement:', 'here is the', 'här är'
    ]
    for line in lines:
        l_lower = line.strip().lower()
        if any(l_lower.startswith(p) for p in skip_prefixes):
            continue
        if (l_lower.startswith('* ') or l_lower.startswith('- ')) and ':' in l_lower:
            if any(p in l_lower for p in skip_prefixes):
                continue
        cleaned.append(line)
    
    result = '\n'.join(cleaned).strip()
    # Strip wrapping quotes if AI added them
    if result.startswith('"') and result.endswith('"'):
        result = result[1:-1].strip()
    return result

# ---------------------------------------------------------
# AI TRANSFORMATION FUNCTIONS WITH FALLBACK
# ---------------------------------------------------------
def transform_image_or_text_with_ai(raw_text, area_choice, uploaded_img=None):
    """Transforms image or raw text into formal Swedish for school documentation."""
    if HAS_GENAI and GEMINI_CONFIGURED:
        model_names = ['gemini-1.5-flash', 'gemini-1.5-flash-latest', 'gemini-2.0-flash', 'gemini-1.5-pro']
        
        prompt = f"""Du är en erfaren rektor och chef i svensk skola.
Omvandla följande råa/handskrivna anteckningar från ett medarbetarsamtal till tjänstemannamässig, formell och professionell svenska lämplig för kommunens skoldokumentation.

Område i samtalet: {area_choice}

KRITISKA INSTRUKTIONER FÖR SVARET:
1. Generera BARA den färdiga formella texten.
2. Generera INGEN analys, inga tankesteg, inga rubriker, inga engelska ord, inga punktlistor, inga citattecken och inga inledande hälsningar.
3. Texten ska vara färdigformaterad för att klistras in direkt i en Word-mall.
4. Tolka handstilen noggrant och formulera allt professionellt på svensk skolsvenska."""

        contents = []
        if uploaded_img is not None:
            try:
                img = Image.open(uploaded_img)
                contents.append(img)
            except Exception as e:
                st.warning(f"Kunde inte läsa bildfilen: {e}")

        if raw_text and raw_text.strip():
            contents.append(f"Råa anteckningar:\n{raw_text}")

        if contents:
            contents.append(prompt)
            for m_name in model_names:
                try:
                    model = genai.GenerativeModel(m_name)
                    response = model.generate_content(contents)
                    if response and response.text:
                        cleaned = clean_ai_output(response.text)
                        if cleaned:
                            return cleaned
                except Exception:
                    continue

    # Rule-based fallback if no AI or AI fails
    text_to_process = raw_text if raw_text else "Medarbetaren bidrar konstruktivt till verksamhetens måluppfyllelse."
    lines = [l.strip() for l in text_to_process.split('.') if l.strip()]
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

def generate_smart_goal_with_ai(criterion, school_goal, raw_idea, raw_actions, followup):
    """Generates a structured SMART goal using AI or fallback rules."""
    if HAS_GENAI and GEMINI_CONFIGURED and raw_idea.strip():
        model_names = ['gemini-1.5-flash', 'gemini-1.5-flash-latest', 'gemini-2.0-flash', 'gemini-1.5-pro']
        prompt = f"""Du är rektor och expert på SMART-modellen i skolan.
Skapa ett konkret SMART-mål (Specifikt, Mätbart, Accepterat, Realistiskt, Tidsatt) baserat på följande uppgifter:

Lönekriterium: {criterion}
Skolans prioriterade mål: {school_goal}
Utvecklingsområde/Idé: {raw_idea}
Åtgärder & Stöd: {raw_actions}
Tidsram: {followup}

Svara exakt i följande tre rader:
MÅL: [Det SMART-formulerade målet]
ÅTGÄRDER: [Åtgärder samt tydlig ansvarsfördelning mellan medarbetare och chef/skola]
UPPFÖLJNING: [Uppföljningsplan och tidsram]"""

        for m_name in model_names:
            try:
                model = genai.GenerativeModel(m_name)
                response = model.generate_content(prompt)
                if response and response.text:
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
                        return clean_ai_output(g_text), clean_ai_output(a_text), clean_ai_output(f_text)
            except Exception:
                continue

    # Fallback rules
    smart_goal_text = f"Utveckla och integrera {raw_idea.lower()} i den dagliga undervisningen för att öka tillgängligheten och studieron för alla elever, kopplat till skolans prioriterade mål om '{school_goal}'."
    smart_action_text = f"Medarbetaren ansvarar för att planera och genomföra åtgärderna ({raw_actions}). Chef/skola ansvarar för resurser, kompetensutveckling samt förutsättningar vid ämneslagsträffar."
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
        
        headers = ["Mål / utvecklingsområde (SMART)", "Åtgärder & Ansvarsfördelning", "Uppföljningsplan"]
        widths = [Inches(2.5), Inches(2.5), Inches(1.5)]
        
        hdr_cells = plan_table.rows[0].cells
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
    
    c_headers = ["Lönekriterium", "Bedömd Nivå", "Motivering & Undervisningsexempel"]
    c_widths = [Inches(1.8), Inches(1.8), Inches(2.9)]
    
    c_hdr_cells = crit_table.rows[0].cells
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
    "Yrkeskunnande/färdighet": "Pedagogisk skicklighet, didaktisk förmåga, anpassning efter elevers behov och variation i undervisningsmetoder.",
    "Engagemang/ansvarsförmåga": "Aktivt ansvar för elevers lärande, utveckling samt deltagande i skolans kvalitetsarbete.",
    "Professionellt förhållningssätt": "Trygghet i yrkesrollen, etiskt bemötande, efterlevnad av styrdokument och värdegrund.",
    "Samarbetsförmåga": "Aktiv delaktighet i kollegialt lärande, samarbete i ämneslag/arbetslag samt dialog med vårdnadshavare.",
    "Ledarskapsförmåga inom yrket": "Förmåga att skapa studiero, trygghet, struktur och en inkluderande lärmiljö i klassrummet.",
    "Flexibilitet/utvecklingsmöjlighet": "Öppenhet för nya arbetssätt, digitalisering, analys av resultat och kontinuerlig yrkesutveckling."
}

FRITIDS_CRITERIA_INFO = {
    "Yrkeskunnande/färdighet": "Fritidspedagogisk skicklighet, stimulerande av elevers utveckling och lärande samt meningsfull fritid.",
    "Engagemang/ansvarsförmåga": "Ansvarsfelexibilitet i hela skoldagen, elevhälsoarbete och delaktighet i fritidshemmets kvalitet.",
    "Professionellt förhållningssätt": "Etiskt förhållningssätt, skapande av trygga relationer och bemötande gentemot elever och vårdnadshavare.",
    "Samarbetsförmåga": "Samverkan mellan fritidshem och skola/förskoleklass samt drivande i kollegialt lärande.",
    "Ledarskapsförmåga inom yrket": "Ledarskap i elevgrupp, konflikthantering, skapande av trygga och tillgängliga lärmiljöer.",
    "Flexibilitet/utvecklingsmöjlighet": "Anpassningsförmåga till skoldagens skiftande behov och utveckling av fritidshemmets uppdrag."
}

# ---------------------------------------------------------
# MAIN STREAMLIT APPLICATION
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

    # TAB 1: OMVANDLA ANTECKNINGAR
    with tab1:
        st.subheader("📝 Omvandla handskrivna/råa anteckningar till officiell text")
        st.info("💡 Ladda upp ett foto på dina handskrivna anteckningar eller klistra in råtext nedan. Appen omvandlar dem till tjänstemannamässig svenska, uppdelat efter kommunens dokumentationsmall, helt utan extra rubriker så att du kan klistra in texten direkt i dina Word-dokument.")
        
        col_input, col_output = st.columns([1, 1])
        
        with col_input:
            st.markdown("### 📥 Dina råa anteckningar / foto")
            section_choice = st.selectbox(
                "Välj område att bearbeta",
                ["Arbetssituation – nuläge", "Resultat & Lönekriterier (Tillbakablick)", "Övriga minnesanteckningar"]
            )
            
            uploaded_img = st.file_uploader(
                "📷 Ladda upp foto på dina anteckningar (PNG, JPG, JPEG):",
                type=["png", "jpg", "jpeg"]
            )
            
            raw_text = st.text_area(
                "Eller klistra in råtext från dina anteckningar:",
                value="Har mkt att göra vid betygssättning. Bra trivsel i laget. Funkar fint m föräldrar. Vill utveckla bildstödet i klassrummet för att nå elever som har svårt m instruktioner. Tränar regelbundet på gymmet.",
                height=140
            )
            
            btn_transform = st.button("⚡ Omvandla till officiell malltext", type="primary", use_container_width=True)
            
        with col_output:
            st.markdown("### 📤 Färdigformaterad text (Redo att klistras in i Word)")
            
            if btn_transform or uploaded_img or raw_text:
                result_text = transform_image_or_text_with_ai(raw_text, section_choice, uploaded_img)
                st.text_area("Kopiera texten nedan (Inga extra rubriker, redo för copy-paste):", value=result_text, height=260)
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

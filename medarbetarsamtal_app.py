import streamlit as st
import io
import re
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

# Custom Styling - Light Theme with High Contrast Dark Text
st.markdown("""
<style>
    /* Force Light Theme for entire app */
    .stApp {
        background-color: #F8FAFC !important;
        color: #0F172A !important;
    }
    
    /* Headers */
    .main-header {
        font-size: 2.2rem;
        color: #1E3A8A !important;
        font-weight: 700;
        margin-bottom: 0.2rem;
    }
    .sub-header {
        font-size: 1.1rem;
        color: #334155 !important;
        margin-bottom: 1.5rem;
    }
    
    /* Sidebar Light Theme & Text */
    [data-testid="stSidebar"] {
        background-color: #F1F5F9 !important;
        border-right: 1px solid #CBD5E1 !important;
    }
    [data-testid="stSidebar"] * {
        color: #0F172A !important;
    }
    
    /* All Input Labels and Text */
    label, p, span, div, h1, h2, h3, h4, h5, h6 {
        color: #0F172A !important;
    }
    
    /* Form inputs and text areas */
    .stTextInput input, .stTextArea textarea, .stSelectbox div[data-baseweb="select"] {
        background-color: #FFFFFF !important;
        color: #000000 !important;
        border: 1px solid #94A3B8 !important;
        border-radius: 6px !important;
    }
    
    /* Ensure scrollbar in text areas */
    textarea {
        color: #000000 !important;
        background-color: #FFFFFF !important;
        font-size: 0.95rem !important;
        line-height: 1.5 !important;
    }
    
    /* Status Badges */
    .status-badge-ok {
        background-color: #DEF7EC;
        color: #03543F !important;
        padding: 4px 10px;
        border-radius: 6px;
        font-weight: 600;
        font-size: 0.85rem;
        display: inline-block;
    }
    .status-badge-info {
        background-color: #E1EFFE;
        color: #1E40AF !important;
        padding: 4px 10px;
        border-radius: 6px;
        font-weight: 600;
        font-size: 0.85rem;
        display: inline-block;
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
# AI HELPER FUNCTIONS & CLEANER
# ---------------------------------------------------------
def clean_ai_output(text):
    """Filters out any internal reasoning, analysis steps, or markdown clutter."""
    if not text:
        return ""
    # Remove markdown code fences if present
    text = re.sub(r'```(?:markdown|text)?\n?', '', text)
    text = text.replace('```', '')
    
    lines = text.split('\n')
    clean_lines = []
    
    for line in lines:
        l_str = line.strip()
        l_lower = l_str.lower()
        
        # Skip meta headers or analysis chatter
        if any(l_lower.startswith(prefix) for prefix in [
            'role:', 'task:', 'input:', 'constraints:', 'topic:', 're-evaluating', 
            'correction/refinement', 'point 1', 'point 2', 'top part', 'main section',
            'wait,', 'let\'s look', 'i will interpret'
        ]):
            continue
        if l_str.startswith('* Role:') or l_str.startswith('* Task:') or l_str.startswith('* Constraints:'):
            continue
        if l_str.startswith('*   Role:') or l_str.startswith('*   Task:') or l_str.startswith('*   Constraints:'):
            continue
        if l_str.startswith('*   *') or l_str.startswith('*   Input:'):
            continue
            
        clean_lines.append(line)
        
    return "\n".join(clean_lines).strip()

def call_gemini_with_fallback(prompt, image=None):
    """Calls Gemini with automatic model fallback to avoid 404 errors."""
    if not (HAS_GENAI and GEMINI_CONFIGURED):
        return None
        
    candidate_models = ['gemini-1.5-flash', 'gemini-1.5-flash-latest', 'gemini-2.0-flash', 'gemini-1.5-pro']
    
    try:
        available = [m.name.replace('models/', '') for m in genai.list_models() if 'generateContent' in m.supported_generation_methods]
        if available:
            flash_models = [m for m in available if 'flash' in m]
            if flash_models:
                candidate_models = flash_models + [m for m in available if m not in flash_models]
            else:
                candidate_models = available + candidate_models
    except Exception:
        pass
        
    for model_name in candidate_models:
        try:
            model = genai.GenerativeModel(model_name)
            if image is not None:
                response = model.generate_content([prompt, image])
            else:
                response = model.generate_content(prompt)
            if response and response.text:
                return response.text.strip()
        except Exception:
            continue
            
    return None

def transform_notes_with_ai(raw_text, area_choice, image_file=None):
    """Transforms raw notes or image of notes into formal Swedish text for documentation."""
    system_prompt = f"""Du är en erfaren rektor och chef i svensk skola.
Omvandla innehållet i följande anteckningar (som kan vara skriven text eller en bild på handskrivna anteckningar) från ett medarbetarsamtal till tjänstemannamässig, formell och professionell svenska lämplig för kommunens skoldokumentation.

Samtalsområde: {area_choice}

KRITISKA REGLER (MÅSTE FÖLJAS EXAKT):
1. Generera ENDAST den slutgiltiga formella texten på svenska.
2. Inga analyssteg, inga tankegångar, inga rubriker, inga punktlistor, inga inledande hälsningar ("Här är...") och inga citationstecken.
3. Ingen engelsk text överhuvudtaget. Svara ENDAST med den färdiga svenskspråkiga brödtexten.
4. Texten ska vara direkt klar att klistras in i en Word-mall.
"""

    pil_img = None
    if image_file is not None:
        try:
            pil_img = Image.open(image_file)
        except Exception:
            pil_img = None

    if pil_img is not None or (raw_text and raw_text.strip()):
        text_to_send = system_prompt
        if raw_text and raw_text.strip():
            text_to_send += f"\nRåa anteckningar:\n{raw_text}"
        if pil_img is not None:
            text_to_send += "\nTolka handstilen i bifogad bild och omvandla alla fakta till formell text."
            
        ai_res = call_gemini_with_fallback(text_to_send, image=pil_img)
        if ai_res:
            return clean_ai_output(ai_res)

    # Fallback rule-based if AI fails or is unconfigured
    if raw_text and raw_text.strip():
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
        
    return ""

def generate_smart_goal_with_ai(criterion, school_goal, raw_idea, raw_actions, followup):
    """Generates a structured SMART goal using AI or fallback rules."""
    if HAS_GENAI and GEMINI_CONFIGURED and raw_idea.strip():
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

        res = call_gemini_with_fallback(prompt)
        if res:
            lines = res.strip().split('\n')
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
    
    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)
    
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
    
    doc.add_paragraph()

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
            
    doc.add_paragraph()

    h1 = doc.add_heading("1. Arbetssituation och Nuläge", level=1)
    h1.style.font.color.rgb = RGBColor(0x1E, 0x3A, 0x8A)
    
    for topic, text in nulage_data.items():
        p_topic = doc.add_paragraph()
        r_t = p_topic.add_run(f"• {topic}: ")
        r_t.bold = True
        p_topic.add_run(text)

    doc.add_paragraph()

    h2 = doc.add_heading("2. Individuell Utvecklingsplan (Mål & Åtgärder)", level=1)
    h2.style.font.color.rgb = RGBColor(0x1E, 0x3A, 0x8A)
    
    if plan_data:
        plan_table = doc.add_table(rows=1, cols=3)
        plan_table.autofit = False
        
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
# MAIN APPLICATION
# ---------------------------------------------------------
def main():
    with st.sidebar:
        st.image("https://img.icons8.com/color/96/school.png", width=70)
        st.title("⚙️ Inställningar")
        
        if GEMINI_CONFIGURED:
            st.markdown('<span class="status-badge-ok">🟢 AI (Gemini) Aktiverad</span>', unsafe_allow_html=True)
        else:
            st.markdown('<span class="status-badge-info">🔵 Regelbaserat läge</span>', unsafe_allow_html=True)

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

    # HEADER
    st.markdown('<div class="main-header">🏫 Medarbetarsamtal & Målgenerator för Skolan</div>', unsafe_allow_html=True)
    st.markdown('<div class="sub-header">Omvandla råa/fotade anteckningar till officiell, färdigformaterad dokumentation och knivskarpa SMART-mål redo för kommunens Word-mallar.</div>', unsafe_allow_html=True)

    # TABS
    tab1, tab2, tab3, tab4 = st.tabs([
        "📝 Omvandla Anteckningar (Word-klar)",
        "🎯 SMART-Målgenerator",
        "📊 Lönekriterier & Bedömning",
        "📄 Förhandsgranska & Exportera Word"
    ])

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
        st.info("💡 Ladda upp ett foto på dina handskrivna anteckningar ELLER klistra in råtext nedan. Appen omvandlar dem till tjänstemannamässig svenska, uppdelat efter kommunens dokumentationsmall, helt utan extra rubriker så att du kan klistra in texten direkt i dina Word-dokument.")
        
        col_input, col_output = st.columns([1, 1])
        
        with col_input:
            st.markdown("### 📥 Dina råa anteckningar / Fototext")
            section_choice = st.selectbox(
                "Välj område i mallen",
                ["Arbetssituation – nuläge", "Resultat & Lönekriterier (Tillbakablick)", "Övriga minnesanteckningar"]
            )
            
            uploaded_image = st.file_uploader(
                "📷 Ladda upp foto på dina handskrivna anteckningar",
                type=["png", "jpg", "jpeg", "webp"],
                help="Släpp en bildfil här eller ta en bild med mobilen."
            )
            
            if uploaded_image is not None:
                st.image(uploaded_image, caption="Uppladdat foto", use_container_width=True)
                
            raw_text = st.text_area(
                "Alternativt: Klistra in text från dina anteckningar:",
                value="",
                placeholder="T.ex. Lugnt i biblioteket, nöjd m arbetslaget, önskar utveckla bildstöd...",
                height=130
            )
            
            btn_transform = st.button("⚡ Omvandla till officiell malltext", type="primary", use_container_width=True)

        with col_output:
            st.markdown("### 📤 Färdigformaterad text (Redo för Word)")
            
            # Compute formatted output text
            result_text = ""
            if btn_transform or uploaded_image or raw_text.strip():
                result_text = transform_notes_with_ai(raw_text, section_choice, uploaded_image)
                
            if not result_text:
                result_text = "Ladda upp ett foto till vänster eller klistra in dina anteckningar och klicka på 'Omvandla till officiell malltext' så visas din färdiga Word-text här."
                
            # THIS TEXTAREA IS ALWAYS RENDERED AND VISIBLE
            st.text_area(
                "Kopiera texten nedan (Inga extra rubriker - redo för copy-paste):",
                value=result_text,
                height=320
            )
            
            if result_text and not result_text.startswith("Ladda upp ett foto"):
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

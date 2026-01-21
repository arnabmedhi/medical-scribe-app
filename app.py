import streamlit as st
from PIL import Image
import j  # Your backend logic
import concurrent.futures
import time
from datetime import datetime # Added for timestamping feedback

# --- PAGE CONFIGURATION (Mobile Optimized) ---
st.set_page_config(
    page_title="AI Medical Scribe",
    page_icon="🏥",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# --- CSS STYLING ---
st.markdown("""
<style>
    /* Mobile-Friendly Header */
    .main-header { 
        font-size: 2rem; 
        color: #2E86C1; 
        text-align: center; 
        margin-bottom: 1rem; 
        font-weight: 700; 
    }
    
    /* Card Style */
    .case-container { 
        background-color: #f8f9fa; 
        padding: 15px; 
        border-radius: 12px; 
        border-left: 6px solid #2E86C1; 
        box-shadow: 0 2px 5px rgba(0,0,0,0.05); 
        margin-bottom: 20px; 
    }
    
    /* Make buttons huge and tappable on mobile */
    .stButton>button { 
        width: 100%; 
        border-radius: 8px; 
        height: 3em; 
        font-weight: 600;
    }
    
    /* Processing Badge */
    @keyframes pulse { 0% { opacity: 0.6; } 50% { opacity: 1; } 100% { opacity: 0.6; } }
    .processing-badge { 
        color: #e67e22; 
        font-weight: bold; 
        padding: 10px; 
        border: 1px solid #e67e22; 
        border-radius: 5px; 
        background-color: #fdf2e9;
        animation: pulse 2s infinite;
        text-align: center;
        margin-bottom: 10px;
    }
    
    /* Links */
    a { text-decoration: none; font-weight: bold; }
    
    /* Feedback Box Style */
    .feedback-box {
        margin-top: 50px;
        padding: 20px;
        border-top: 1px solid #ddd;
        background-color: #f1f8ff;
        border-radius: 10px;
    }
</style>
""", unsafe_allow_html=True)

# --- SESSION STATE SETUP ---
if 'page' not in st.session_state: st.session_state.page = 'home'
if 'cases' not in st.session_state: st.session_state.cases = [0] 

# Stores RESULT DATA
if 'results' not in st.session_state: st.session_state.results = {} 

# Stores ACTIVE JOBS
if 'active_jobs' not in st.session_state: st.session_state.active_jobs = {}

# Background Worker Pool
if 'executor' not in st.session_state:
    st.session_state.executor = concurrent.futures.ThreadPoolExecutor(max_workers=4)

# --- HELPER FUNCTIONS ---
def go_home(): st.session_state.page = 'home'
def go_medicine(): st.session_state.page = 'medicine'

def add_case():
    new_id = (max(st.session_state.cases) + 1) if st.session_state.cases else 0
    st.session_state.cases.append(new_id)

def remove_case(case_id):
    if case_id in st.session_state.cases: st.session_state.cases.remove(case_id)
    if case_id in st.session_state.results: del st.session_state.results[case_id]
    if case_id in st.session_state.active_jobs: del st.session_state.active_jobs[case_id]
    st.rerun()

def save_feedback(text):
    """Saves user feedback to a local text file with a timestamp."""
    try:
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with open("feedback_log.txt", "a") as f:
            f.write(f"[{timestamp}] {text}\n" + "-"*30 + "\n")
        return True
    except Exception as e:
        return False

# THE BACKGROUND TASK
def background_task(images, model):
    try:
        # Returns: {'link':..., 'id':..., 'name':...}
        return j.run_pipeline(images, model_choice=model)
    except Exception as e:
        return {"error": str(e)}

# --- STATUS MONITOR FRAGMENT ---
@st.fragment(run_every=2)
def status_monitor(case_id):
    # 1. Check if Job is Active
    if case_id in st.session_state.active_jobs:
        future = st.session_state.active_jobs[case_id]
        if future.done():
            # Job finished!
            data = future.result()
            
            if "error" in data:
                 st.session_state.results[case_id] = {"error": data['error']}
            else:
                # ⚡ FETCH FILE BYTES ⚡
                file_bytes = j.export_docx(data['id'])
                
                st.session_state.results[case_id] = {
                    "link": data['link'],
                    "name": data['name'],
                    "bytes": file_bytes
                }
            
            del st.session_state.active_jobs[case_id]
            st.rerun()
        else:
            st.markdown(f"<div class='processing-badge'>⏳ AI is working on Case {case_id}...</div>", unsafe_allow_html=True)
    
    # 2. Check if Job is Done (Show Result)
    elif case_id in st.session_state.results:
        res = st.session_state.results[case_id]
        
        if "error" in res:
            st.error(res['error'])
        else:
            st.success("✅ Ready!")
            
            c1, c2 = st.columns([1, 1])
            with c1:
                st.markdown(f"📄 **[Preview in Browser]({res['link']})**")
                st.caption("Opens in Google Docs")

            with c2:
                if res['bytes']:
                    st.download_button(
                        label="⬇️ Download Word Doc",
                        data=res['bytes'],
                        file_name=f"{res['name']}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        key=f"dl_{case_id}",
                        use_container_width=True
                    )
                else:
                    st.warning("Download unavailable")

            st.info("ℹ️ **Tip:** Download the file above to save it locally.")

        if st.button("🔄 Start Over", key=f"restart_{case_id}"):
            del st.session_state.results[case_id]
            st.rerun()

# =========================================================
# PAGE 1: HOME
# =========================================================
if st.session_state.page == 'home':
    st.markdown("<div class='main-header'>🏥 AI Discharge Tool</div>", unsafe_allow_html=True)
    
    col1, col2 = st.columns(2)
    with col1:
        st.info("💊 **Internal Medicine**")
        if st.button("Enter Medicine ➡️", use_container_width=True):
            go_medicine(); st.rerun()

    with col2:
        st.warning("🔪 **Surgery**")
        st.button("Locked 🔒", disabled=True, use_container_width=True)

# =========================================================
# PAGE 2: MEDICINE DASHBOARD
# =========================================================
elif st.session_state.page == 'medicine':
    
    c1, c2 = st.columns([1, 6])
    with c1: 
        if st.button("⬅️ Home"): go_home(); st.rerun()
    with c2: st.markdown("### 💊 Internal Medicine")

    # --- RENDER CASES ---
    for case_id in st.session_state.cases:
        with st.container():
            st.markdown(f"<div class='case-container'>", unsafe_allow_html=True)
            
            h1, h2 = st.columns([8, 2])
            with h1: st.markdown(f"**📂 Case ID: {case_id}**")
            with h2: 
                if st.button("🗑️", key=f"del_{case_id}", help="Delete Session"):
                    remove_case(case_id)

            status_monitor(case_id)

            if case_id not in st.session_state.results and case_id not in st.session_state.active_jobs:
                uploaded_files = st.file_uploader(f"Upload Notes", type=["jpg","png","jpeg"], key=f"up_{case_id}", accept_multiple_files=True)
                
                st.write("") 
                model = st.radio("Select Intelligence:", ("Gemini 2.5 Flash (Fast)", "Gemini 2.5 Pro (Best)"), index=1, key=f"mod_{case_id}")
                model_clean = "Gemini 2.5 Flash" if "Flash" in model else "Gemini 2.5 Pro"

                if st.button(f"⚡ Process Case {case_id}", key=f"btn_{case_id}", type="primary"):
                    if uploaded_files:
                        pil_images = [Image.open(f) for f in uploaded_files]
                        future = st.session_state.executor.submit(background_task, pil_images, model_clean)
                        st.session_state.active_jobs[case_id] = future
                        st.rerun()
                    else:
                        st.warning("⚠️ Upload images first")

            st.markdown("</div>", unsafe_allow_html=True)

    # ADD BUTTON
    if st.button("➕ New Patient Case"):
        add_case()
        st.rerun()

    # --- FEEDBACK SECTION (NEW) ---
    st.write("---")
    with st.expander("💬 Report an Issue / Send Feedback"): #
        with st.form("feedback_form"):
            user_feedback = st.text_area("Tell us what's wrong or suggest a feature:", placeholder="Type here...") #
            submitted = st.form_submit_button("Submit Feedback") #
            
            if submitted and user_feedback:
                if save_feedback(user_feedback):
                    st.success("✅ Thank you! Your feedback has been saved.")
                else:
                    st.error("❌ Error saving feedback. Check permissions.")
import os
import streamlit as st
import concurrent.futures
import time
from PIL import Image
from datetime import datetime
import j           # <-- Backend for Medicine
import j_surgery   # <-- Backend for Surgery (NEW)

# --- AUTHENTICATION FIREWALL ---
if "GOOGLE_TOKEN" in st.secrets:
    secret_value = st.secrets["GOOGLE_TOKEN"]
    # Write the file and force the OS to save it immediately
    with open("token.json", "w") as f:
        f.write(secret_value)
        f.flush()
        os.fsync(f.fileno()) # <--- This forces the write instantly
    
if "OPENAI_KEY" in st.secrets:
    os.environ["OPENAI_API_KEY"] = st.secrets["OPENAI_KEY"]

# --- PAGE CONFIGURATION ---
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
    
    /* Surgery Card Style (Red Border) */
    .surgery-container {
        background-color: #fff5f5;
        padding: 15px; 
        border-radius: 12px; 
        border-left: 6px solid #c0392b; 
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
    
    a { text-decoration: none; font-weight: bold; }
</style>
""", unsafe_allow_html=True)

# --- SESSION STATE SETUP ---
if 'page' not in st.session_state: st.session_state.page = 'home'
if 'cases' not in st.session_state: st.session_state.cases = [0] 
if 'results' not in st.session_state: st.session_state.results = {} 
if 'active_jobs' not in st.session_state: st.session_state.active_jobs = {}

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
    try:
        return j.save_feedback_online(text)
    except Exception as e:
        return False

# THE BACKGROUND TASK (GENERIC)
# We pass the 'module' (j or j_surgery) to this function now
def background_task(images, model, backend_module):
    try:
        # Call run_pipeline on the specific backend (j or j_surgery)
        return backend_module.run_pipeline(images, model_choice=model)
    except Exception as e:
        return {"error": str(e)}

# --- STATUS MONITOR FRAGMENT ---
@st.fragment(run_every=2)
def status_monitor(case_id):
    if case_id in st.session_state.active_jobs:
        future = st.session_state.active_jobs[case_id]
        if future.done():
            data = future.result()
            
            if isinstance(data, str):
                st.session_state.results[case_id] = {"error": data}
            elif "error" in data:
                 st.session_state.results[case_id] = {"error": data['error']}
            else:
                # Success!
                # We use j.export_docx because the file ID is universal on Drive
                file_bytes = j.export_docx(data['id'])
                
                cost_info = data.get('cost', "N/A")
                if cost_info != "N/A":
                    st.toast(f"💰 Cost: {cost_info}")

                st.session_state.results[case_id] = {
                    "link": data['link'],
                    "name": data['name'],
                    "bytes": file_bytes,
                    "cost": cost_info 
                }
            
            del st.session_state.active_jobs[case_id]
            st.rerun()
        else:
            st.markdown(f"<div class='processing-badge'>⏳ AI is working on Case {case_id}...</div>", unsafe_allow_html=True)
    
    elif case_id in st.session_state.results:
        res = st.session_state.results[case_id]
        
        if "error" in res:
            st.error(res['error'])
        else:
            st.success("✅ Ready!")
            
            c1, c2 = st.columns([1, 1])
            with c1:
                st.markdown(f"📄 **[Preview]({res['link']})**")
            with c2:
                if res['bytes']:
                    st.download_button(
                        label="⬇️ Download Doc",
                        data=res['bytes'],
                        file_name=f"{res['name']}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        key=f"dl_{case_id}",
                        use_container_width=True
                    )
                else:
                    st.warning("Download unavailable")

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
        st.error("🔪 **General Surgery**") # Changed to Red for distinction
        # UNLOCKED: Now navigates to 'surgery' page
        if st.button("Enter Surgery ➡️", use_container_width=True):
            st.session_state.page = 'surgery'
            st.rerun()

# =========================================================
# PAGE 2: MEDICINE DASHBOARD
# =========================================================
elif st.session_state.page == 'medicine':
    
    c1, c2 = st.columns([1, 6])
    with c1: 
        if st.button("⬅️ Home"): go_home(); st.rerun()
    with c2: st.markdown("### 💊 Internal Medicine")

    for case_id in st.session_state.cases:
        with st.container():
            st.markdown(f"<div class='case-container'>", unsafe_allow_html=True)
            
            h1, h2 = st.columns([8, 2])
            with h1: st.markdown(f"**📂 Case ID: {case_id}**")
            with h2: 
                if st.button("🗑️", key=f"del_{case_id}"): remove_case(case_id)

            status_monitor(case_id)

            if case_id not in st.session_state.results and case_id not in st.session_state.active_jobs:
                uploaded_files = st.file_uploader(f"Upload Notes", type=["jpg","png","jpeg"], key=f"up_{case_id}", accept_multiple_files=True)
                
                st.write("") 
                model = st.radio("Select Intelligence:", ("Gemini 2.5 Flash (Fast)", "Gemini 2.5 Pro (Best)"), index=1, key=f"mod_{case_id}")
                model_clean = "Gemini 2.5 Flash" if "Flash" in model else "Gemini 2.5 Pro"

                if st.button(f"⚡ Process Medicine", key=f"btn_{case_id}", type="primary"):
                    if uploaded_files:
                        pil_images = [Image.open(f) for f in uploaded_files]
                        # CALLS 'j' (MEDICINE BACKEND)
                        future = st.session_state.executor.submit(background_task, pil_images, model_clean, j)
                        st.session_state.active_jobs[case_id] = future
                        st.rerun()
                    else:
                        st.warning("⚠️ Upload images first")

            st.markdown("</div>", unsafe_allow_html=True)

    if st.button("➕ New Patient Case"):
        add_case()
        st.rerun()
        
    st.write("---")
    with st.expander("💬 Report an Issue / Send Feedback"):
        with st.form("feedback_form_med"):
            user_feedback = st.text_area("Tell us what's wrong:", height=150)
            submitted = st.form_submit_button("Submit")
            if submitted and user_feedback:
                if save_feedback(user_feedback): st.success("✅ Saved.")
                else: st.error("❌ Error.")

# =========================================================
# PAGE 3: SURGERY DASHBOARD (NEW)
# =========================================================
elif st.session_state.page == 'surgery':
    
    c1, c2 = st.columns([1, 6])
    with c1: 
        if st.button("⬅️ Home"): go_home(); st.rerun()
    with c2: st.markdown("### 🔪 General Surgery")

    for case_id in st.session_state.cases:
        with st.container():
            # Use distinct red styling for surgery cards
            st.markdown(f"<div class='surgery-container'>", unsafe_allow_html=True)
            
            h1, h2 = st.columns([8, 2])
            with h1: st.markdown(f"**📂 Surgery Case: {case_id}**")
            with h2: 
                # Unique key 'del_s_' to avoid conflict with medicine buttons
                if st.button("🗑️", key=f"del_s_{case_id}"): remove_case(case_id)

            status_monitor(case_id)

            if case_id not in st.session_state.results and case_id not in st.session_state.active_jobs:
                # Unique key 'up_s_'
                uploaded_files = st.file_uploader(f"Upload Surgery Notes", type=["jpg","png","jpeg"], key=f"up_s_{case_id}", accept_multiple_files=True)
                
                st.write("") 
                # Unique key 'mod_s_'
                model = st.radio("Select Intelligence:", ("Gemini 2.5 Flash", "Gemini 2.5 Pro"), index=1, key=f"mod_s_{case_id}")
                model_clean = "Gemini 2.5 Flash" if "Flash" in model else "Gemini 2.5 Pro"

                # Unique key 'btn_s_'
                if st.button(f"⚡ Generate Surgery Discharge", key=f"btn_s_{case_id}", type="primary"):
                    if uploaded_files:
                        pil_images = [Image.open(f) for f in uploaded_files]
                        
                        # --- CRITICAL CHANGE: CALLS 'j_surgery' BACKEND ---
                        future = st.session_state.executor.submit(background_task, pil_images, model_clean, j_surgery)
                        
                        st.session_state.active_jobs[case_id] = future
                        st.rerun()
                    else:
                        st.warning("⚠️ Upload images first")

            st.markdown("</div>", unsafe_allow_html=True)

    if st.button("➕ New Surgery Case"):
        add_case()
        st.rerun()

    st.write("---")
    with st.expander("💬 Surgery Feedback"):
        with st.form("feedback_form_surg"):
            user_feedback = st.text_area("Tell us what's wrong:", height=150)
            submitted = st.form_submit_button("Submit")
            # ... inside the Surgery Page Feedback Form ...
    if submitted and user_feedback:
        # We add this tag so you know it came from Surgery
        tagged_feedback = f"🔴 [SURGERY DEPT]: {user_feedback}" 
        
        if save_feedback(tagged_feedback): 
            st.success("✅ Saved to central feedback folder.")
        else: 

            st.error("❌ Error.")

import io # Required for handling image bytes
import time
from googleapiclient.http import MediaIoBaseUpload # Required for uploading files to Drive

import warnings
warnings.filterwarnings("ignore")


import socket
socket.setdefaulttimeout(600)  # Force Windows to wait 600 seconds (10 mins) before failing

import google.generativeai as genai
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
import json
import os
import re  # The Nuclear Cleaning Tool
from dotenv import load_dotenv  # Loads the secret .env file

import csv
import datetime
from pathlib import Path

# --- PRICING TABLE (Jan 2026) ---
# We define the cost per 1 Million tokens for your models
PRICING = {
    "gemini-2.5-pro":   {"input": 1.25, "output": 10.00},
    "gemini-2.5-flash": {"input": 0.30, "output": 2.50},
    "gemini-3-flash-preview": {"input": 0.50, "output": 3.00},
    "gemini-3-pro-preview":   {"input": 2.00, "output": 12.00}
}

def log_usage(response, model_name, note=""):
    try:
        # 1. Get Token Counts
        usage = response.usage_metadata
        in_tokens = usage.prompt_token_count
        out_tokens = usage.candidates_token_count
        
        # 2. Calculate Cost
        rates = PRICING.get(model_name, {"input": 0, "output": 0})
        total_cost = ((in_tokens / 1_000_000) * rates["input"]) + ((out_tokens / 1_000_000) * rates["output"])
        
        cost_string = f"₹{total_cost * 87:.2f}"
        
        # 3. Create Log String
        log_details = (
            f"Date: {datetime.datetime.now()}\n"
            f"Model: {model_name}\n"
            f"Cost: {cost_string}\n"
        )
        
        print(f"   [💰 Cost Logged: {cost_string} for {note}]")
        
        # --- CRITICAL FIX: You MUST return these two values ---
        return cost_string, log_details 
        # ------------------------------------------------------

    except Exception as e:
        print(f"   [⚠️ Cost Logging Failed: {e}]")
        # In case of error, return defaults so it doesn't crash
        return "N/A", ""

# ==============================================================================
# SECTION A: CONFIGURATION (Loaded from .env file)
# ==============================================================================
load_dotenv()

GENAI_API_KEY = os.getenv("GENAI_API_KEY")

# --- NEW: Pointing to OBGYN variables ---
MASTER_TEMPLATE_ID = os.getenv("OBS_TEMPLATE_ID")      # <--- NEW
OUTPUT_FOLDER_ID = os.getenv("OBS_OUTPUT_FOLDER_ID")   # <--- NEW
IMAGES_FOLDER_ID = os.getenv("OBS_IMAGES_FOLDER_ID")   # <--- NEW

# 2. SHARED FOLDERS (Same as Surgery/Medicine)
CLIENT_SECRET_FILE = os.getenv("CLIENT_SECRET_FILE")
COST_FOLDER_ID = os.getenv("COST_FOLDER_ID")           # <--- Shared
FEEDBACK_FOLDER_ID = os.getenv("FEEDBACK_FOLDER_ID")   # <--- Shared


# 3. Safety Check
if not GENAI_API_KEY or not MASTER_TEMPLATE_ID:
    print("CRITICAL ERROR: Keys are missing.")
    print("Make sure you created the '.env' file with your API keys.")
    exit()

# ==============================================================================
# SECTION B.4: OBSTETRICS LAB TEST ORDER
# ==============================================================================
# This list controls the VERTICAL order of rows in your Google Doc Table.
# The script will fill cells in exactly this sequence under each date.
# ==============================================================================
# SECTION B.3: OBSTETRICS LAB TEST ORDER
# ==============================================================================
# This must match the exact rows in your Google Doc Table vertically
OBS_TEST_ORDER = [
    "hb",                  # Row 1
    "tlc",                 # Row 2
    "plt",                 # Row 3
    "pt_inr",              # Row 4 (Combined)
    "bil_total_direct",    # Row 5 (Combined)
    "sgpt_sgot_alp",       # Row 6 (Combined)
    "hiv_hbsag_rpr_hcv",   # Row 7 (Combined)
    "urea_creat",          # Row 8 (Combined)
    "na_k_cl_ca",          # Row 9 (Combined)
    "fbs_ppbs",            # Row 10 (Combined)
    "tsh",                 # Row 11
    "t3_t4",               # Row 12 (Combined)
    "urine_pus_epi_rbc",   # Row 13 (Combined)
    "urine_cs"             # Row 14
]



# ==============================================================================
# SECTION C: THE PLACEHOLDERS (Matches your Google Doc)
# ==============================================================================
placeholder_rules = {
    # --- BASIC INFO ---
    "{{patient_name}}": "Extract Full Name",
    "{{uhid}}": "Extract UHID/Registration Number",
    "{{age}}": "Extract Age only (e.g. 66 YRS)",
    "{{address}}": "Extract complete address",
    "{{aadhar}}": "Extract AADHAR No. If empty, leave blank.",
    "{{unit}}": "Extract Unit Number, only number (e.g. 3).",
    "{{unit_incharge}}": "Extract names listed under Unit In-charge or Unit Head (e.g. Dr. Farhanul Huda, Dr. Sudhir Kumar Singh).",
    # --- SECTION A: OBSTETRICS HEADER & MANAGEMENT ---
    
    "{{lscs_date}}": (
        "STRICTLY extract the 'Date of LSCS' exactly as written, including the TIME. "
        "Example: '20/1/26 2:00 AM'"
    ),

    "{{management}}": (
        "Extract the 'Management' section details. "
        "**STRICT FORMATTING RULE:** You must place every distinct category on a separate new line. Do not combine them into a paragraph.\n"
        "**Required Output Structure:**\n"
        "[Procedure Name and Indication]\n"
        "Done by: [List of Surgeons]\n"
        "Anesthesia: [Type of Anesthesia]\n"
        "Anesthetist: [Name of Anesthetist]"
    ),

    "{{diagnosis}}": (
        "Extract the 'Diagnosis' exactly as written. "
        "Include the Obstetric Formula (Primigravida/GPLA) and Period of Gestation (POG).\n"
        "Example: 'Primigravida at 40+1 wks POG'"
    ),

    # ---  ADMISSION DETAILS ---
    "{{doa}}": "Date of Admission (DOA)",
    "{{dod}}": "Date of Discharge (DOD)",
    "{{consultant}}": "Extract 'Consultant in-charge' name specifically.",
    
    # --- OBSTETRIC HISTORY SECTIONS (Strict Detailed Rules) ---

    "{{brief_history}}": (
        "Write the 'Brief History' paragraph in formal medical documentation style. "
        "**STRICT OUTPUT RULES:**\n"
        "1. **Format:** 4-5 lines, concise, objective, past tense. Do not summarize; include every specific detail found.\n"
        "2. **Content Order:** Start with Booking Status (Booked/Unbooked) + Gravida/Parity + Gestational Age (POG) with basis (LMP and/or Scan) -> Presenting reason to LR/OPD -> Current Admission Status -> Key Negative Symptoms (LPV/BPV/pain abdomen) -> Fetal Movements (if mentioned).\n"
        "3. **Abbreviations:** Use standard terms like G/P, POG, LR, i/v/o, c/o, LPV, BPV.\n"
        "4. **Accuracy:** Keep all numbers, dates, and scan names exactly as written in the source text.\n"
        "5. **Completeness:** If other relevant details (e.g., referral source, specific complaints) are present, include them."
    ),

    "{{anc_history_t1}}": (
        "Write the 'Antenatal History – T1 (First Trimester)' section. "
        "**STRICT OUTPUT RULES:**\n"
        "1. **Format:** Crisp bullet points or short sentences (4-8 lines). No summarization of key medical facts.\n"
        "2. **Content to Extract:** Mode of conception + Confirmation method (UPT/Scan) -> Folic Acid intake -> Key Negatives (LPV/BPV/Pain Abdomen) -> Fever history (with/without rash) -> Nausea/Vomiting severity -> Radiation exposure -> Teratogenic drug intake.\n"
        "3. **Style:** Preserve 'No h/o...' style for negatives. Remove repetition but keep medically standard grammar.\n"
        "4. **Completeness:** If specific medications or other T1 events are mentioned, they MUST be included."
    ),

    "{{anc_history_t2_t3}}": (
        "Write the 'Antenatal History – T2/T3 (Second and Third Trimester)' section. "
        "**STRICT OUTPUT RULES:**\n"
        "1. **Format:** 7-10 lines, structured and objective.\n"
        "2. **Content Order:** Quickening (Month/POG) -> Iron & Calcium intake -> Tetanus (Td) doses -> Key Negatives (LPV/BPV/Pain Abdomen) -> History of Hypertension/Diabetes/Thyroid issues -> LFT/ICP symptoms (itching palms/soles) -> Parenteral Iron / Blood Transfusion history.\n"
        "3. **Abbreviations:** Use standard abbreviations: Td, BP, RBS/BS, LFT, ICP.\n"
        "4. **Completeness:** Include any other specific complications or medications mentioned in the T2/T3 period."
    ),

    # --- MENSTRUAL & OBSTETRIC HISTORY ---

    "{{lmp}}": (
        "STRICTLY extract the 'Last Menstrual Period' (LMP) date exactly as written. "
        "Example: '14/4/25'."
    ),

    "{{pmc}}": (
        "Extract 'Past Menstrual Cycles' details VERBATIM. "
        "Include regularity, cycle length (days), flow duration, and presence of dysmenorrhea/clots or any other relevant details and symptoms "
        "Example: 'Regular cycles/28-30 days/ 4-5 days/ not a/w dysmenorrhea'."
    ),

    "{{obs_history}}": (
        "Extract 'O/H' (Obstetric History) and 'ML' (Married Life) details. "
        "**STRICT FORMAT:** Combine them into clear lines.\n"
        "1. Married Life (ML): Extract duration (e.g., 'ML = 3 years').\n"
        "2. Obstetric Score: Extract status (e.g., 'Primigravida', 'G2P1L1') exactly as written."
    ),

    # --- PAST & SURGICAL HISTORY (Logic: Specific Data vs "Not Significant") ---

    "{{past_history}}": (
        "Extract 'Past History'. "
        "**LOGIC RULE:**\n"
        "1. If specific diseases (HTN, TB, Diabetes, Asthma, etc.) are mentioned, LIST THEM.\n"
        "2. If the text says 'Not significant', 'Nil', 'NAD', or is empty, OUTPUT EXACTLY: 'Not significant'."
    ),

    "{{family_history}}": (
        "Extract 'Family History'. "
        "**LOGIC RULE:**\n"
        "1. If family history of HTN, DM, Twins, or congenital anomalies is mentioned, LIST THEM.\n"
        "2. If the text says 'Not significant', 'Nil', 'NAD', or is empty, OUTPUT EXACTLY: 'Not significant'."
    ),

    "{{surgical_history}}": (
        "Extract 'Surgical history'. "
        "**LOGIC RULE:**\n"
        "1. If any previous surgery (Appendectomy, Laparoscopy, etc.) is mentioned, LIST IT with the year if available.\n"
        "2. If the text says 'Not significant', 'Nil', 'NAD', or is empty, OUTPUT EXACTLY: 'Not significant'."
    ),
   
    
    # --- EXAM (GENERAL) - THE +/- VE LOGIC ---
    "{{pallor}}": "Look for Pallor. If present -> 'Present'. If absent/normal -> 'Absent'.",
    "{{icterus}}": "Look for Icterus. If present -> 'Present'. If absent/normal -> 'Absent'.",
    "{{cyanosis}}": "Look for Cyanosis. If present -> 'Present'. If absent/normal -> 'Absent'.",
    "{{clubbing}}": "Look for Clubbing. If present -> 'Present'. If absent/normal -> 'Absent'.",
    "{{lymphadenopathy}}": "Look for Lymphadenopathy. If present -> 'Present'. If absent/normal -> 'Absent'.",
    "{{edema}}": "Look for Pedal Edema. If present -> 'Present'. If absent/normal -> 'Absent'.",
    # --- VITALS ---
    "{{gc}}": (
        "Extract General Condition (GC). "
        "**SMART DEFAULT:** If the notes say 'Normal', 'GC Fair', or nothing specific, "
        "OUTPUT EXACTLY: 'Conscious, well oriented to time, place & person, alert and cooperative.' "
        "If the notes specify something else (e.g., 'Drowsy', 'Disoriented', 'Poor'), extract that text verbatim."
    ),
    "{{pulse}}": "Extract Pulse Rate number only (e.g. 101).",
    "{{rr}}": "Extract Respiratory Rate (RR) number only.",
    "{{bp}}": "Extract Blood Pressure (BP) number only",
    "{{temp}}": "Extract Temperature, number only. If no data present than put it as afebrile",
    "{{spo2}}": "Extract SpO2 percentage, number only",
    "{{Height}}": "Extract Height, number only",
    "{{Weight}}": "Extract Weight, number only",

    # --- SYSTEMIC EXAM (SMART DEFAULTS) ---
    "{{cvs_exam}}": "CVS Findings. If normal, use default 'S1, S2 present'. If abnormal, describe findings.",
    "{{rs_exam}}": "Respiratory Findings. If normal, use default 'B/L Air entry present'.",

    # --- OBSTETRIC EXAMINATION (Strict Concise Format) ---

    "{{pa_exam}}": (
        "Generate the 'P/A (Per Abdomen)' examination finding in concise OBG case-sheet style. "
        "**OUTPUT RULES:**\n"
        "- Output ONLY one line.\n"
        "- Keep comma-separated phrases (no full sentences).\n"
        "- Include ONLY what is present: uterine size/term size, tone, lie/presentation, contractions/relaxed, FHR with bpm and regularity, tenderness.\n"
        "- Use standard abbreviations exactly: FHR, bpm, /R (regular) or /I (irregular).\n"
        "- Do NOT add interpretation or extra findings.\n"
        "Example: 'Uterus term size, Soft, cephalic, relaxed, FHR-155bpm/R.'"
    ),

    "{{le_exam}}": (
        "Generate the 'L/E (Local Examination)' finding in concise OBG case-sheet style. "
        "**OUTPUT RULES:**\n"
        "- Output ONLY one line.\n"
        "- If notes indicate normal, write 'NAD' or 'Normal'.\n"
        "- If abnormal findings are present (edema, discharge, bleeding, lesions), list them in short comma-separated terms exactly as given.\n"
        "- Do NOT add interpretation or missing details."
    ),

    "{{pv_exam}}": (
        "Generate the 'P/V (Per Vaginal)' examination finding in concise OBG case-sheet style. "
        "**OUTPUT RULES:**\n"
        "- Output ONLY one line.\n"
        "- Use the same compact, comma-separated format.\n"
        "- Include at least: cervix consistency/position/dilatation (fingers or cm), effacement, station, membrane status, pelvis adequacy, and any procedure done (e.g., sweep/stretch).\n"
        "- Keep numeric values and wording exactly as given (e.g., '2 finger loose', 'station – 0').\n"
        "- Do NOT add Bishop score, contractions, or assumed details."
    ),

    # --- BLOOD GROUP ---
    "{{blood_group}}": "STRICTLY extract the 'Blood group' exactly as written. Example: 'B positive'.",

    # --- OBSTETRICS LAB GRID (Complex Slash Logic) ---
    # --- OBSTETRICS LAB GRID (Granular Extraction - NO SLASHES) ---
    # --- OBSTETRICS LAB GRID (The Strict Recipe) ---
    # --- OBSTETRICS LAB GRID (Granular Extraction - NO SLASHES) ---
    # --- OBSTETRICS LAB GRID (Standard Slash Logic) ---
    "{{labs_json}}": (
        f"EXTRACT LABS as JSON for keys: {', '.join(OBS_TEST_ORDER)}. \n"
        "**CRITICAL COMBINATION RULE:** You MUST join multiple values with a slash ' / '. \n"
        "If a specific value is missing, use '-'.\n\n"
        "**EXAMPLES:**\n"
        "- 'pt_inr': PT 13.5, INR 1.1 -> output '13.5 / 1.1'\n"
        "- 'bil_total_direct': Total 1.2, Direct 0.4 -> output '1.2 / 0.4'\n"
        "- 'sgpt_sgot_alp': SGPT 40, SGOT 35, ALP 110 -> output '40 / 35 / 110'\n"
        "- 'hiv_hbsag_rpr_hcv': HIV NR, HBsAg Neg -> output 'NR / Neg / - / -'\n"
        "- 'urea_creat': Urea 24, Creat 0.9 -> output '24 / 0.9'\n"
        "- 'na_k_cl_ca': Na 136, K 4.2 -> output '136 / 4.2 / - / -'\n"
        "- 'fbs_ppbs': FBS 90, PPBS 140 -> output '90 / 140'\n"
        "- 't3_t4': T3 1.2, T4 9.8 -> output '1.2 / 9.8'\n"
        "- 'urine_pus_epi_rbc': Pus 2-3, Epi 4-5, RBC Nil -> output '2-3 / 4-5 / Nil'\n"
        "Format: {'12/02/2026': {'hb': '10.1', 'pt_inr': '12.8 / 1.01'}, ...}"
    ),

    # --- DYNAMIC LABS (HPLC & PERIPHERAL SMEAR - Max 4) ---
    "{{hplc_smear_json}}": (
        "EXTRACT HPLC and PERIPHERAL SMEAR data as a JSON Object.\n"
        "**RULES:**\n"
        "1. **Full Text:** Extract the RESULT exactly as written in the document. Do not summarize. Include all sentences. Do not miss any word.\n"
        "2. **Sort Chronologically:** Earliest date is index 0 (Date 1).\n"
        "3. **Structure:** {'hplc': [{'date': 'DD/MM/YY', 'result': 'Full verbatim text...'}], 'ps': [{'date': 'DD/MM/YY', 'result': 'Full verbatim text...'}]}\n"
        "4. **Max:** Extract ALL available reports (up to 4 each)."
    ),

    # --- DYNAMIC USG SERIES (Max 7) ---
    "{{usg_series_json}}": (
        "EXTRACT USG OBSTETRICS reports as a JSON List.\n"
        "**RULES:**\n"
        "1. **Full Text:** Extract the RESULT exactly as written. Transcribe SLIUF, POG, FHR, Liquor, Impression, etc. verbatim.\n"
        "2. **Sort Chronologically:** Earliest date first.\n"
        "3. **Structure:** Return a list: [{'date': 'DD/MM/YY', 'result': 'Full verbatim text...'}, ...]\n"
        "4. **Max:** Extract ALL available reports (up to 7)."
    ),

    "{{hospital_course}}": (
        "Output ONLY the section starting exactly with: 'Course during hospital stay:'. "
        "Convert raw notes into 3–6 short, medically standard lines (or one compact paragraph). "
        "Strictly transcribe ONLY what is present: admission reason, investigations, surveillance, "
        "key decision for procedure (indication), timing (elective/emergency). "
        "Keep abbreviations standard (i/v/o, LSCS, PROM, CTG, USG)."
    ),

    "{{per_op_findings}}": (
        "Output ONLY the section starting exactly with: 'Per op findings:'. "
        "Format as BULLET POINTS or short step-wise lines. "
        "Strictly include ONLY what is written: incision type, uterine segment, uterine incision, "
        "liquor, presentation, neonatal status, delayed cord clamping, placenta/membranes, "
        "uterus closure, hemostasis, counts, sheath/skin closure, and drugs given (e.g., 'Tab. Misoprostol 800 mg PR')."
    ),

    # --- POST-OP & ADVICE ---
    "{{post_op_course}}": (
        "Output ONLY the section starting exactly with: 'Post op course:'. "
        "Format as BULLET POINTS or short step-wise lines. "
        "Write 5–10 lines in chronological order (POD-wise). "
        "Strictly include ONLY what is present: antibiotics/analgesics duration, catheter/SRC removal, "
        "voiding, bowel movements, ambulation, wound status (discharge/induration/erythema), "
        "contraception counseling, and mother–baby status at discharge."
    ),

    "{{discharge_advice}}": (
        "Output ONLY the section starting exactly with: 'Advice on discharge:'. "
        "Format as CLEAN BULLET POINTS. "
        "Strictly transcribe: activity restrictions, wound care, medicines (Name + Dose + Route + Freq + Duration/SOS), "
        "breastfeeding, immunization, diet, follow-up, and suture removal date. "
        "Standardize format (e.g., 'Tab. Taxim 200 mg PO BD x 5 days') but DO NOT change the drug/dose."
    ),

    # --- BABY DETAILS (Strict Extraction) ---

    # --- BABY DETAILS (Raw Data Only) ---

    "{{sex_of_baby}}": (
        "Output ONLY the sex of the baby exactly as written in the notes.\n"
        "**OUTPUT RULES:**\n"
        "- Do NOT start with 'Sex of baby:'. Just give the value.\n"
        "- Example: 'Female' or 'Male (Boy)'.\n"
        "- Do NOT add any extra words."
    ),

    "{{birth_date_time}}": (
        "Output ONLY the Date and Time of birth exactly as written.\n"
        "**OUTPUT RULES:**\n"
        "- Do NOT start with 'Date:'. Just give the value.\n"
        "- Preserve the exact format found (e.g., '20/01/2026 @ 12:36 AM')."
    ),

    "{{birth_weight}}": (
        "Output ONLY the Birth Weight exactly as written.\n"
        "**OUTPUT RULES:**\n"
        "- Do NOT start with 'Birth weight:'. Just give the number and unit.\n"
        "- Example: '2990 gm' or '2.5 kg'.\n"
        "- Do NOT add any interpretation."
    ),

    "{{apgar_score}}": (
        "Extract the APGAR Score values exactly as found in the text. "
        "Look for numbers associated with '1 min' and '5 min'. "
        "Return the raw text found (e.g., '8 at 1 min, 9 at 5 min' or '7/10, 8/10'). "
        "Do not format it. Just copy the text."
    ),
    # --- STAFF DETAILS (Strict Vertical List) ---
    "{{junior_residents}}": (
        "Extract names of JUNIOR RESIDENTS. "
        "**STRICT FORMAT:** List each name on a NEW LINE. Do not use commas. "
        "Example:\nDr. Anshuman\nDr. Nandakumar\nDr. Rudra"
    ),

    "{{senior_residents}}": (
        "Extract names of SENIOR RESIDENTS. "
        "**STRICT FORMAT:** List each name on a NEW LINE. Do not use commas. "
        "Example:\nDr. Ashish Mishra\nDr. Shahid"
    ),

    "{{consultants}}": (
        "Extract names of CONSULTANTS. "
        "**STRICT FORMAT:** List each name on a NEW LINE. Do not use commas. "
        "Example:\nDr. Farhanul Huda\nDr. Sudhir Kumar Singh"
    ),

}


# ==============================================================================
# SECTION D: THE LOGIC ENGINE
# ==============================================================================

def get_user_credentials():
    """Handles Google Login with Aggressive Retries for Cloud Stability."""
    SCOPES = ['https://www.googleapis.com/auth/documents', 'https://www.googleapis.com/auth/drive']
    creds = None
    
    # --- PHASE 1: AGGRESSIVE RETRY LOOP (Max 5 Seconds) ---
    # We try 5 times to grab the token.json in case it is being written or is locked.
    for attempt in range(5):
        if os.path.exists('token.json'):
            try:
                creds = Credentials.from_authorized_user_file('token.json', SCOPES)
                if creds and creds.valid:
                    # Found it! Stop waiting and return immediately.
                    return creds
                if creds and creds.expired and creds.refresh_token:
                    creds.refresh(Request())
                    return creds
            except Exception as e:
                print(f"⚠️ Attempt {attempt+1}: token.json found but read failed ({e}). Retrying...")
        else:
            print(f"⚠️ Attempt {attempt+1}: token.json not found yet. Retrying...")
        
        # Wait 1 second before trying again (unless it's the last attempt)
        if attempt < 4:
            time.sleep(1)

    # --- PHASE 2: IF TOKEN STILL FAILS AFTER 5 TRIES ---
    
    # CRITICAL CHECK: Only try browser login if we are LOCALLY running (File exists).
    # This prevents the "No such file" crash on Streamlit Cloud.
    if CLIENT_SECRET_FILE and os.path.exists(CLIENT_SECRET_FILE):
        print("--- Launching Browser for Local Login... ---")
        flow = InstalledAppFlow.from_client_secrets_file(CLIENT_SECRET_FILE, SCOPES)
        creds = flow.run_local_server(port=0)
        
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
        return creds
    
    else:
        # If we are on Cloud and Token failed 5 times, STOP HERE.
        # Do not crash. Just raise a clear error.
        raise Exception("CRITICAL ERROR: Could not load 'token.json' after 5 attempts. The Google Token in Secrets is missing or invalid.")


def normalize_key(text):
    """Turns 'EsR', 'esr', 'ESR', 'Es_r' into just 'esr' for easier matching."""
    if not text: return ""
    return str(text).lower().replace("_", "").replace(" ", "").replace("-", "").strip()

def fill_smart_grid(service, doc_id, labs_data, test_order, anchor_name):
    """Fills grid starting EXACTLY at the anchor column."""
    print(f"   -> Processing Grid for Anchor: {anchor_name}...")
    doc = service.documents().get(documentId=doc_id).execute(num_retries=5)
    content = doc.get('body').get('content')

    anchor_found = False
    table_index = -1
    anchor_row = -1
    anchor_col = -1
    
    clean_anchor_text = anchor_name.replace("{{", "").replace("}}", "").strip()

    # 1. Find the Anchor
    for i, element in enumerate(content):
        if 'table' in element:
            for r_idx, row in enumerate(element['table']['tableRows']):
                for c_idx, cell in enumerate(row['tableCells']):
                    full_cell_text = ""
                    if 'content' in cell:
                        for content_item in cell['content']:
                            if 'paragraph' in content_item:
                                for elem in content_item['paragraph']['elements']:
                                    if 'textRun' in elem:
                                        full_cell_text += elem['textRun']['content']
                    
                    if clean_anchor_text in full_cell_text:
                        anchor_found = True
                        table_index = i
                        anchor_row = r_idx
                        anchor_col = c_idx
                        print(f"      FOUND {anchor_name} at Table {i}, Row {r_idx}, Col {c_idx}")
                        break
                if anchor_found: break
        if anchor_found: break
    
    if not anchor_found:
        print(f"      WARNING: {anchor_name} NOT FOUND.")
        return []

    requests = []
    # 2. Clean the Anchor Text (So it doesn't stay behind)
    requests.append({
        'replaceAllText': {
            'containsText': {'text': anchor_name, 'matchCase': True},
            'replaceText': ' ' 
        }
    })

    # 3. Sort Dates
    sorted_dates = sorted(labs_data.keys())
    table = content[table_index]['table']

    # 4. Loop through Dates
    for i, date in enumerate(sorted_dates):
        # --- THE FIX IS HERE ---
        # We use 'anchor_col + i' so the first date (i=0) goes exactly into the Anchor Cell.
        target_col = anchor_col + i 
        
        if target_col >= len(table['tableRows'][0]['tableCells']): break

        try:
            # A. Fill Date Header (Top Row)
            cell = table['tableRows'][anchor_row]['tableCells'][target_col]
            end_index = cell['endIndex'] - 1
            requests.append({'insertText': {'location': {'index': end_index}, 'text': str(date)}})
            
            # B. Fill Test Values
            day_map = {normalize_key(k): v for k, v in labs_data[date].items()}
            
            for test_idx, python_key in enumerate(test_order):
                target_row = anchor_row + 1 + test_idx
                
                if target_row >= len(table['tableRows']): break
                
                search_key = normalize_key(python_key)
                val = str(day_map.get(search_key, ""))
                
                # Cleanup
                val = val.replace("{", "").replace("}", "").replace("[", "").replace("]", "").strip()

                if val and val != "":
                    try:
                        cell = table['tableRows'][target_row]['tableCells'][target_col]
                        end_index = cell['endIndex'] - 1
                        requests.append({'insertText': {'location': {'index': end_index}, 'text': val}})
                    except: pass
        except: continue

    # 5. Sort requests
    requests.sort(key=lambda x: x.get('insertText', {}).get('location', {}).get('index', 0), reverse=True)
    return requests

# --- NEW: Upload Images to a Specific Folder ---
def upload_patient_images(image_list, patient_name):
    if not IMAGES_FOLDER_ID: return 
    
    creds = get_user_credentials()
    drive_service = build('drive', 'v3', credentials=creds)
    
    try:
        # 1. Create a Sub-Folder for this specific Patient
        folder_metadata = {
            'name': f"{patient_name} - Images",
            'parents': [IMAGES_FOLDER_ID],
            'mimeType': 'application/vnd.google-apps.folder'
        }
        folder = drive_service.files().create(body=folder_metadata, fields='id').execute()
        patient_folder_id = folder.get('id')
        
        # 2. Upload all images into that sub-folder
        print(f"   -> Backing up {len(image_list)} images to Drive...")
        for i, img in enumerate(image_list):
            img_byte_arr = io.BytesIO()
            img.save(img_byte_arr, format='JPEG')
            img_byte_arr.seek(0)
            
            file_metadata = {
                'name': f"Page_{i+1}.jpg",
                'parents': [patient_folder_id]
            }
            media = MediaIoBaseUpload(img_byte_arr, mimetype='image/jpeg')
            drive_service.files().create(body=file_metadata, media_body=media).execute()
            
    except Exception as e:
        print(f"   ⚠️ Image Backup Failed: {e}")

# --- NEW: Save Cost Log to Drive ---
def log_cost_to_drive(text_content, patient_name):
    if not COST_FOLDER_ID: return
    
    creds = get_user_credentials()
    drive_service = build('drive', 'v3', credentials=creds)
    docs_service = build('docs', 'v1', credentials=creds)
    
    try:
        # Create a simple Google Doc for the cost log
        title = f"Cost - {patient_name} - {datetime.datetime.now().strftime('%d-%m')}"
        file_metadata = {
            'name': title,
            'parents': [COST_FOLDER_ID],
            'mimeType': 'application/vnd.google-apps.document'
        }
        doc = drive_service.files().create(body=file_metadata).execute()
        
        # Write the cost details
        requests = [{'insertText': {'location': {'index': 1}, 'text': text_content}}]
        docs_service.documents().batchUpdate(documentId=doc.get('id'), body={'requests': requests}).execute()
        print("   -> Cost Log Saved to Drive.")
    except Exception as e:
        print(f"   ⚠️ Cost Drive Upload Failed: {e}")

# --- NEW: Save Feedback to Drive ---
def save_feedback_online(text):
    if not FEEDBACK_FOLDER_ID: return False
    
    creds = get_user_credentials()
    drive_service = build('drive', 'v3', credentials=creds)
    docs_service = build('docs', 'v1', credentials=creds)
    
    try:
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
        file_metadata = {
            'name': f"Feedback - {timestamp}",
            'parents': [FEEDBACK_FOLDER_ID],
            'mimeType': 'application/vnd.google-apps.document'
        }
        doc = drive_service.files().create(body=file_metadata).execute()
        
        requests = [{'insertText': {'location': {'index': 1}, 'text': text}}]
        docs_service.documents().batchUpdate(documentId=doc.get('id'), body={'requests': requests}).execute()
        return True
    except Exception as e:
        print(f"Feedback Error: {e}")
        return False

def run_pipeline(image_list=None, model_choice="Gemini 2.5 Pro"):  # <--- 1. Accept model choice
    
    # 1. Validation
    if not image_list:
        return "Error: No images provided to Logic Engine."
    
    print("--- 1. Authenticating... ---")
    creds = get_user_credentials()
    
    # 2. Map User Choice to Actual Model ID
    # This connects the Frontend Radio Button to the Backend Logic
    model_map = {
        "Gemini 2.5 Flash": "models/gemini-2.5-flash",
        "Gemini 2.5 Pro": "models/gemini-2.5-pro",
        "Gemini 3.0 Pro": "models/gemini-3-pro-preview" # Added just in case
    }
    
    # Default to Pro if something goes wrong
    selected_model_id = model_map.get(model_choice, "models/gemini-2.5-pro")
    
    print(f"--- 2. AI Processing using {model_choice} ({selected_model_id}) ---")
    
    genai.configure(api_key=GENAI_API_KEY)
    
    model = genai.GenerativeModel(
        selected_model_id,  # <--- 3. Use the selected model here
        safety_settings=[
            {"category": "HARM_CATEGORY_HARASSMENT", "threshold": "BLOCK_NONE"},
            {"category": "HARM_CATEGORY_HATE_SPEECH", "threshold": "BLOCK_NONE"},
            {"category": "HARM_CATEGORY_SEXUALLY_EXPLICIT", "threshold": "BLOCK_NONE"},
            {"category": "HARM_CATEGORY_DANGEROUS_CONTENT", "threshold": "BLOCK_NONE"},
        ]
    )

    # --- YOUR CUSTOM PROMPT (MERGED WITH NEW LAB RULES) ---
    final_prompt_text = f"""
    You are a professional medical scribe. Extract data from the handwritten note below.

    You are an expert Medical Scribe. 
    I have provided {len(image_list)} images of handwritten patient notes.
    
    ### CRITICAL INSTRUCTION: DEEP READING MODE
    The handwriting is messy. Do NOT jump to the JSON immediately.

    - Key: "{{labs_json}}". Keys: {OBS_TEST_ORDER} 

    ### 0. LOGIC FOR LAB DATA (NEW):
    - Extract ALL lab values associated with dates.
    - Return a JSON object under key "{{labs_json}}".
    # CHANGED: Uses OBS_TEST_ORDER now
    - Keys must match EXACTLY: {OBS_TEST_ORDER} 
    

    ### CRITICAL INSTRUCTION: SEMANTIC SEARCH
    Understand the **biological meaning** of tests. 
    - Map "CK-MB" or "Creatine Kinase-MB" to "cpkmb".
    - Map "GeneXpert" or "MTB/RIF" to "csfcbnaat".
    - Map "Cell Count" in CSF to "csftlc".


    ### 1. LOGIC FOR "+ve / -ve" FIELDS:
    For placeholders {{pallor}}, {{icterus}}, {{cyanosis}}, {{clubbing}}, {{lymphadenopathy}}, {{edema}}:
    - READ the text carefully.
    - If the note says "Present", "Positive", "++" -> Output: "Present"
    - If the note says "Absent", "Negative", "--", "Nil", or DOES NOT MENTION it -> Output: "Absent"


    ### 3. RULES FOR RESIDENTS/FACULTY:
    - Extract exact titles like "(SR): DR NAME", "(JR): DR NAME". 
    - Keep them on separate lines.



    ### 6. DATA CLEANING:
    - Return ONLY valid JSON.
    - Keys must match the placeholders exactly (e.g., "{{pallor}}").
    - Values must be PLAIN TEXT. Do NOT include brackets {{ }} in the values.

    ### INPUT DATA:
    PLACEHOLDERS: {json.dumps(placeholder_rules)}
    """
    
    # --- COMBINE PROMPT + IMAGES ---
    prompt_content = [final_prompt_text]
    prompt_content.extend(image_list)

    # Initialize variables
    cost_display = "N/A"
    log_content = ""

    try:
        # Run AI
        response = model.generate_content(prompt_content)

        clean_model_name = selected_model_id.replace("models/", "")
        
        # --- NEW: Capture Cost Data ---
        cost_display, log_content = log_usage(response, clean_model_name, note=f"Medical Extraction ({model_choice})")
        
        # Clean JSON
        raw_text = response.text
        start_index = raw_text.find('{')
        end_index = raw_text.rfind('}') + 1
        
        if start_index != -1 and end_index != -1:
                json_text = raw_text[start_index:end_index]
                extracted_data = json.loads(json_text)

                # =========================================================
                # LOGIC 1: UNPACK HPLC & PERIPHERAL SMEAR (Scope-Safe)
                # =========================================================
                # 1. Initialize Empty First (Prevents Crash if AI skips key)
                safe_hplc_data = {} 

                # 2. Try to fill it
                if "{{hplc_smear_json}}" in extracted_data:
                    val = extracted_data["{{hplc_smear_json}}"]
                    if isinstance(val, str): 
                        try: safe_hplc_data = json.loads(val)
                        except: safe_hplc_data = {}
                    elif isinstance(val, dict):
                        safe_hplc_data = val
                    
                    # Remove key so it doesn't print
                    del extracted_data["{{hplc_smear_json}}"]

                # 3. Handle HPLC (4 slots) - Uses safe_hplc_data
                hplc_list = safe_hplc_data.get("hplc", [])
                for i in range(1, 5): # 1 to 4
                    key_date = f"{{{{hplc_date_{i}}}}}"
                    key_res = f"{{{{hplc_res_{i}}}}}"
                    
                    if i <= len(hplc_list):
                        item = hplc_list[i-1]
                        raw_date = item.get("date", "").strip()
                        extracted_data[key_date] = f"HPLC ({raw_date})" if raw_date else "HPLC"
                        extracted_data[key_res] = item.get("result", "")
                    else:
                        extracted_data[key_date] = ""
                        extracted_data[key_res] = ""

                # 4. Handle Peripheral Smear (4 slots)
                ps_list = safe_hplc_data.get("ps", [])
                for i in range(1, 5): # 1 to 4
                    key_date = f"{{{{ps_date_{i}}}}}"
                    key_res = f"{{{{ps_res_{i}}}}}"
                    
                    if i <= len(ps_list):
                        item = ps_list[i-1]
                        raw_date = item.get("date", "").strip()
                        extracted_data[key_date] = f"Peripheral Smear ({raw_date})" if raw_date else "Peripheral Smear"
                        extracted_data[key_res] = item.get("result", "")
                    else:
                        extracted_data[key_date] = ""
                        extracted_data[key_res] = ""

                # =========================================================
                # LOGIC 2: UNPACK USG SERIES (Scope-Safe)
                # =========================================================
                # 1. Initialize Empty First
                safe_usg_list = []

                # 2. Try to fill it
                if "{{usg_series_json}}" in extracted_data:
                    val = extracted_data["{{usg_series_json}}"]
                    if isinstance(val, str): 
                        try: safe_usg_list = json.loads(val)
                        except: safe_usg_list = []
                    elif isinstance(val, list):
                        safe_usg_list = val
                    
                    del extracted_data["{{usg_series_json}}"]

                # 3. Loop 7 times (Safe)
                for i in range(1, 8): 
                    key_date = f"{{{{usg_date_{i}}}}}"
                    key_res = f"{{{{usg_res_{i}}}}}"
                    
                    if i <= len(safe_usg_list):
                        item = safe_usg_list[i-1]
                        raw_date = item.get("date", "").strip()
                        extracted_data[key_date] = f"USG OBS ({raw_date})" if raw_date else "USG OBS"
                        extracted_data[key_res] = item.get("result", "")
                    else:
                        extracted_data[key_date] = ""
                        extracted_data[key_res] = ""
                
                if "{{usg_series_json}}" in extracted_data: del extracted_data["{{usg_series_json}}"]
                # Check if the AI returned the grids as "Strings" instead of "Objects"
                grid_keys = ["{{labs_json}}", "{{cardiac_json}}", "{{csf_json}}"]
                
                for key in grid_keys:
                    if key in extracted_data:
                        val = extracted_data[key]
                        
                        # If it is a string (The Error Cause), try to fix it
                        if isinstance(val, str):
                            # Clean up common AI mistakes
                            clean_val = val.replace("```json", "").replace("```", "").strip()
                            # Fix single quotes to double quotes
                            if clean_val.startswith("{") and "'" in clean_val:
                                clean_val = clean_val.replace("'", '"')
                            
                            try:
                                print(f"   -> Auto-Correcting stringified JSON for {key}...")
                                extracted_data[key] = json.loads(clean_val)
                            except:
                                # If it's really broken, just make it empty so we don't crash
                                extracted_data[key] = {}
                # --- ROBUST FIX END ---

        else:
                return "Error: AI failed to generate JSON. Try again."
            
    except Exception as e:
        return f"AI Logic Error: {e}"

    # Determine filename
    patient_name = extracted_data.get("{{patient_name}}", "Unknown")
    patient_name = re.sub(r'[\\/*?:"<>|]', "", patient_name)
    
    # --- NEW: Upload Images & Cost Log ---
    if image_list:
        upload_patient_images(image_list, patient_name)
    
    if log_content:
        log_cost_to_drive(log_content, patient_name)
    
    # NEW: Add model tag to filename
    model_tag = "Flash" if "Flash" in model_choice else "Pro"
    new_filename = f"Discharge Summary - {patient_name} ({model_tag})"

    print(f"--- 3. Creating File: '{new_filename}' ---")
    drive_service = build('drive', 'v3', credentials=creds)

    file_metadata = {
        'name': new_filename,
        'parents': [OUTPUT_FOLDER_ID],
        'mimeType': 'application/vnd.google-apps.document'
    }
    
    NEW_DOCUMENT_ID = None
    
    # --- NETWORK RETRY LOOP (The Fix for 10060 Error) ---
    # We try 3 times. If internet flickers, it won't crash.
    for attempt in range(1, 4):
        try:
            print(f"   -> Attempt {attempt} to create file...")
            copy_response = drive_service.files().copy(
                fileId=MASTER_TEMPLATE_ID,
                body=file_metadata,
                supportsAllDrives=True
            ).execute(num_retries=5) 
            
            NEW_DOCUMENT_ID = copy_response.get('id')
            print(f"   --> SUCCESS! File Created ID: {NEW_DOCUMENT_ID}")
            break # It worked! Exit the loop.
            
        except Exception as e:
            print(f"   ⚠️ Attempt {attempt} Failed: {e}")
            time.sleep(3) # Wait 3 seconds and try again

    # If it FAILED all 3 times, then we stop.
    if not NEW_DOCUMENT_ID:
        return {"error": "Network Error: Internet is too slow or blocking Google Drive. Try disabling VPN/Firewall."}

    print(f"--- 4. Filling Data... ---")
    docs_service = build('docs', 'v1', credentials=creds)

    if "{{labs_json}}" in extracted_data:
        lab_data = extracted_data["{{labs_json}}"]
        
        # --- FIX 1: FORCE CONVERT STRING TO DICT ---
        if isinstance(lab_data, str):
            try:
                print(f"   ⚠️ Lab data was a String. Converting to Dict...")
                lab_data = json.loads(lab_data)
            except Exception as e:
                print(f"   ❌ JSON Conversion Failed: {e}")
                lab_data = {}
        # -------------------------------------------

        # Check if it is valid now
        if isinstance(lab_data, dict):
            # Debug Print: Show what keys we found vs what we expect
            print(f"   -> AI Found {len(lab_data)} dates in Labs.")
            
            lab_requests = fill_smart_grid(
                docs_service, 
                NEW_DOCUMENT_ID, 
                lab_data, 
                OBS_TEST_ORDER,    # Ensure this list is defined at the top of obs.py
                "{{LAB_ANCHOR}}"   # Ensure Doc cell has ONLY this text
            )
            
            if lab_requests:
                docs_service.documents().batchUpdate(
                    documentId=NEW_DOCUMENT_ID, 
                    body={'requests': lab_requests}
                ).execute(num_retries=5)
                print("   ✅ OBGYN Lab Grid Filled Successfully.")
            else:
                print("   ⚠️ Lab Grid Logic ran, but generated NO requests. (Check Anchor or Row Count)")
        else:
            print(f"   ❌ Skipping Labs: Data is still {type(lab_data)} (Not a Dict)")
        
        # Cleanup
        if "{{labs_json}}" in extracted_data: del extracted_data["{{labs_json}}"]


    # --- B. FILL TEXT FIELDS (With Smart Defaults Logic) ---
    print("   -> Merging extracted data with Professional Defaults...")
    
    requests = []
    final_data = extracted_data.copy()
    

    # 2. Create Replacement Requests (The "Brute Force" Method)
    for placeholder, value in final_data.items():
        if isinstance(value, dict) or isinstance(value, list): continue # Skip JSON grids
        
        # Clean brackets if any
        text_value = str(value).replace("{", "").replace("}", "").replace("[", "").replace("]", "").strip()

        # --- THE NUCLEAR CLEANER: REMOVES "NONE" and "NOT FOUND" ---
        if text_value.lower() in ["none", "null", "not_found", "not found"]:
            text_value = ""

        req = {
            'replaceAllText': {
                'containsText': {'text': placeholder, 'matchCase': True},
                'replaceText': text_value
            }
        }
        requests.append(req)

    if requests:
        docs_service.documents().batchUpdate(documentId=NEW_DOCUMENT_ID, body={'requests': requests}).execute(num_retries=5)
        print("--- 5. SUCCESS! ---")
        
        # CHANGE 3: Return the link string instead of just printing it
        final_link = f"https://docs.google.com/document/d/{NEW_DOCUMENT_ID}"
    print(f"Link: {final_link}")
    
    # Return all the info we need for the button
    return {
        "link": final_link, 
        "id": NEW_DOCUMENT_ID, 
        "name": new_filename,
        "cost": cost_display # <--- Send cost to App
    }

if __name__ == "__main__":
    run_pipeline()

def export_docx(file_id):
    """Downloads the Google Doc as a .docx file for the user."""
    creds = get_user_credentials() # Uses your existing auth
    drive_service = build('drive', 'v3', credentials=creds)
    try:
        # Request to export the file as a Word Doc
        request = drive_service.files().export_media(
            fileId=file_id,
            mimeType='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )
        file_data = request.execute() # Returns the actual file bytes
        return file_data
    except Exception as e:
        print(f"Export Error: {e}")
        return None
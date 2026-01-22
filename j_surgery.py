import io # Required for handling image bytes
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

# 2. Get keys safely
GENAI_API_KEY = os.getenv("GENAI_API_KEY")
MASTER_TEMPLATE_ID = os.getenv("SURGERY_TEMPLATE_ID")  # <--- CHANGED

# 2. The Folder to save Surgery Discharge Summaries
OUTPUT_FOLDER_ID = os.getenv("SURGERY_OUTPUT_FOLDER_ID") # <--- CHANGED

# 3. The Folder to backup Surgery Patient Images
IMAGES_FOLDER_ID = os.getenv("SURGERY_IMAGES_FOLDER_ID") # <--- CHANGED

CLIENT_SECRET_FILE = os.getenv("CLIENT_SECRET_FILE")
COST_FOLDER_ID = os.getenv("COST_FOLDER_ID")         # Folder to save Cost Logs
FEEDBACK_FOLDER_ID = os.getenv("FEEDBACK_FOLDER_ID") # Folder to save Feedback

# 3. Safety Check
if not GENAI_API_KEY or not MASTER_TEMPLATE_ID:
    print("CRITICAL ERROR: Keys are missing.")
    print("Make sure you created the '.env' file with your API keys.")
    exit()

# ==============================================================================


# ==============================================================================
# SECTION B.2: LAB TEST ORDER (Must match your Table Rows exactly)
# ==============================================================================
 # ==============================================================================
# SECTION B.3: SURGERY LAB TEST ORDER
# ==============================================================================
# This matches the specific rows in your Surgery Discharge Summary Table
SURGERY_TEST_ORDER = [
    "hb", 
    "tlc", 
    "plt", 
    "tb_db",        # Combined Bilirubin
    "sgot_sgpt",    # Combined Liver Enzymes
    "alp_ggt",      # Combined ALP/GGT
    "protein", 
    "albumin", 
    "viral_1",      # First Viral Markers row
    "urea", 
    "cr", 
    "na_k",         # Combined Electrolytes
    "pt_inr",       # Combined PT/INR
    "viral_2",      # Second Viral Markers row (if present)
    "hba1c", 
    "thyroid"       # T3/T4/TSH
]

# 2. CARDIAC / ACUTE PHASE GRID
# These must match the semantic keys in the prompt
CARDIAC_TEST_ORDER = [
    "hstropi", "cpkmb", "cpknac", "esr", "crp", "ldh", 
    "il6", "cortisol", "procal", "bnp"
]

# 3. CSF / FLUID ANALYSIS GRID
# These must match the semantic keys in the prompt
CSF_TEST_ORDER = [
    "tlc", "dlc", "glucose", "protein", "ada", 
    "cbnaat", "gram", "culture", "koh", "india"
]



# ==============================================================================
# SECTION C: THE PLACEHOLDERS (Matches your Google Doc)
# ==============================================================================
placeholder_rules = {
    # --- BASIC INFO ---
    "{{patient_name}}": "Extract Full Name",
    "{{uhid}}": "Extract UHID/Registration Number",
    "{{age}}": "Extract Age only (e.g. 66 YRS)",
    "{{gender}}": "Extract Gender",
    "{{address}}": "Extract complete address",
    "{{aadhar}}": "Extract AADHAR No. If empty, leave blank.",
    "{{contact}}": "Extract Contact no.",
    "{{dop}}": "Extract D.O.P (Date of Procedure/Surgery).",
    "{{unit}}": "Extract Unit Number, only number (e.g. 3).",
    "{{unit_incharge}}": "Extract names listed under Unit In-charge or Unit Head (e.g. Dr. Farhanul Huda, Dr. Sudhir Kumar Singh).",
    

    # ---  ADMISSION DETAILS ---
    "{{doa}}": "Date of Admission (DOA)",
    "{{dod}}": "Date of Discharge (DOD)",
    "{{dop}}": "Extract D.O.P (Date of Procedure/Surgery).",
    "{{consultant}}": "Extract 'Consultant in-charge' name specifically.",
    "{{procedure}}": "Extract the specific Procedure performed (e.g. RIGHT MODIFIED RADICAL MASTECTOMY).",
    
    # --- DIAGNOSIS & SUMMARY ---
    "{{final_diagnosis}}": "Extract Final Diagnosis completely. Maintain medical capitalization.Do not List Diagnosis serially (1., 2., 3.). DO NOT use + signs. **FORBIDDEN CHARACTER:** Do NOT use the + sign. BAD: Dengue +ve, T2DM + HTN. GOOD:** Dengue Positive, Type 2 Diabetes Mellitus and Hypertension.",
   
    # --- 3. CLINICAL HISTORY (Detailed) ---
    "{{complaint}}": "Extract CHIEF COMPLAINT with duration (e.g. LUMP IN RIGHT BREAST x 1 YEAR).",
    
   "{{hpi}}": (
        "Extract HISTORY OF PRESENTING ILLNESS in 3 distinct parts:\n"
        "1. **Narrative:** Start with 'Patient was apparently well...'. Describe onset, duration, progression, and associated symptoms.\n"
        "2. **Positive Symptoms:** Start a new line with '**H/O**'. List ALL present symptoms found in the text.\n"
        "3. **Negative Symptoms:** Start a new line with '**No H/O**'. List ALL denied/absent symptoms found in the text.\n"
        "**SMART CHECKLIST:**\n"
        "- **General Check:** Always check status of [pain, fever, headache, seizure, abdominal distension, jaundice, hemoptysis, chronic cough, bone fractures, loss of appetite, weight loss].\n"
        "- **Breast Cancer Specific:** ONLY IF the case is Breast CA, also check [nipple discharge, ulceration, nipple retraction, skin changes].\n"
        "- **Dynamic Capture:** If the text mentions ANY OTHER symptoms not listed above (e.g., vomiting, constipation, bleeding), you MUST classify them as H/O or No H/O."
    ),

    "{{past_medical}}": (
        "Extract 'Past Medical history'.\n"
        "**SMART CHECKLIST:** Check status of [HTN, T2DM, COPD, Asthma, TB, Thyroid].\n"
        "- If Present: Write 'H/O [Condition] x [Duration]'.\n"
        "- If Absent/Denied: You MUST explicitly write 'No H/O [Condition]'.\n"
        "- Example: 'H/O HTN. No H/O T2DM, COPD, Asthma.'"
    ),

    "{{past_surgical}}": (
        "Extract 'Past surgical history'.\n"
        "- If surgeries are mentioned: List them (e.g., 'H/O Lap Cholecystectomy 5 years back').\n"
        "- **CRITICAL:** If NO surgeries are mentioned, output exactly: 'No significant history'."
    ),

    "{{personal_history}}": (
        "Extract 'Personal History' with Gender Logic.\n"
        "1. **Basics:** Diet (Veg/Mixed), Marital Status.\n"
        "2. **Female Only:** Extract Parity (e.g., P3L3) and Breastfeeding history. IF MALE, DO NOT INCLUDE THIS.\n"
        "3. **Habits (Sleep/Appetite/Bowel/Bladder):**\n"
        "   - If ALL are normal, write: 'Normal sleep, appetite, bowel, and bladder habits.'\n"
        "   - If ANY is abnormal (e.g., Insomnia), write ONLY the abnormal one and add 'rest normal'. (e.g., 'Sleep disturbed, rest normal')."
    ),

    "{{menstrual_history}}": (
        "Extract 'Menstrual History' **ONLY IF FEMALE**.\n"
        "- Include: Menopause status (duration/cause) or LMP.\n"
        "- IF MALE: Return an empty string."
    ),

    "{{family_history}}": (
        "Extract 'Family history' focused on the patient's diagnosis.\n"
        "- If Patient has Cancer or other whatever diagnosis: Explicitly look for family history of similar cancers or other whatever diagnosis(Breast, Ovarian, etc.).\n"
        "- If Present: Detail the relation.\n"
        "- If Absent: Write 'No family history of [Condition] in 1st or 2nd degree relatives'."
    ),

    "{{treatment_history}}": (
        "Extract Treatment/NACT History. "
        "List the specific number of cycles, drug names, and dosages line-by-line and in different  "
        "**STRICT FORMAT:** You must list every item on a new line. Do not combine them with commas. "
        "Example Output:\n"
        "6 cycles of NACT\n"
        "6 cycles of Inj Docetaxel 100 mg\n"
        "Inj Carboplatin 450 mg\n"
        "Inj Trastuzumab 450 mg\n"
        "with 2 cycles of Inj Trastuzumab 380 mg"
    ),

    
    "{{case_summary}}": "Extract Case Summary. Include past history, presenting complaints, and duration.""Write a PROFESSIONAL NARRATIVE. "
        "Start with: 'Name is a Age-year-old Gender, resident of Address...' "
        "Include comorbidities, family history, and presenting complaints numbered serially.",
    
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
    "{{temp}}": "Extract Temperature, number only",
    "{{spo2}}": "Extract SpO2 percentage, number only",
    "{{Height}}": "Extract Height, number only",
    "{{Weight}}": "Extract Weight, number only",

    # --- SYSTEMIC EXAM (SMART DEFAULTS) ---
    "{{cvs_exam}}": "CVS Findings. If normal, use default 'S1 S2 heard, no murmurs'. If abnormal, describe findings.",
    "{{rs_exam}}": "Respiratory Findings. If normal, use default 'B/L normal vesicular breath sounds heard'.",
    "{{pa_exam}}": "Per Abdomen Findings. If normal, use default 'Non distended, non-tender, bowel sounds are present, and no palpable organomegaly Present.'.",
    "{{cns_exam}}": "CNS Findings. If normal, use default 'No focal neurological deficits'.",

    # --- 4. EXAMINATION (Dynamic Local & Smart P/A) ---
    "{{local_exam}}": (
        "Extract 'Local Examination' verbatim but enforce STRICT SUBHEADINGS. "
        "Every anatomical part mentioned must have its own line followed by a colon.\n"
        "**Rule:** Even if a side is normal (NAD), you must write the Heading first, then the finding.\n"
        "**Example Output Structure:**\n"
        "RIGHT BREAST:\n"
        "8x6cm lump, hard, fixed...\n"
        "RIGHT AXILLA:\n"
        "No lymph nodes palpable\n"
        "LEFT BREAST:\n"
        "NAD / No lump palpable\n"
        "LEFT AXILLA:\n"
        "NAD\n"
        "P/A:\n"
        "Soft, non-tender..."
    ),

    "{{pa_inspection}}": (
        "Extract the entire 'Local Examination' section verbatim. "
        "**FORMATTING RULE:** You must preserve the specific sub-headings used in the text . "
        "Write the findings line-by-line under their respective sub-headings. "
        "Do not summarize; capture specific details."
    ),

    "{{pa_palpation}}": (
        "Extract Palpation findings. "
        "**SMART DEFAULT:** If the text says 'Soft', 'Normal', or 'NAD', "
        "OUTPUT EXACTLY: 'Soft, no local rise in temperature, no tenderness, no organomegaly, no lump palpable.' "
        "**EXCEPTION:** If 'Tenderness', 'Guarding', or 'Rigidity' or anything else is mentioned, REPLACE the default with those findings."
    ),

    "{{pa_percussion}}": (
        "Extract Percussion findings. "
        "**SMART DEFAULT:** 'Tympanic note present.' "
        "Change only if findings like 'Dullness' or 'Fluid thrill' or anything else are mentioned."
    ),

    "{{pa_auscultation}}": (
        "Extract Auscultation findings. "
        "**SMART DEFAULT:** 'Bowel sounds present.' "
        "Change only if findings like 'Absent bowel sounds' or 'Increased bowel sounds' or anything else are mentioned."
    ),

    # --- MRI BRAIN (Detailed & Verbatim) ---
    "{{mri_date}}": "Extract the Date of the MRI Brain (e.g., 09/08/25).",

    "{{mri_findings}}": (
        "Extract the 'Imaging Findings' section of the MRI Brain **VERBATIM** (Exact Text). "
        "**DO NOT SUMMARIZE.** "
        "Include every bullet point, every sentence. "
        "Preserve the original structure and formatting."
    ),

    "{{mri_impression}}": (
        "Extract the 'Impression' section of the MRI Brain exactly as written. "
        "Example: 'No evidence of cerebral metastasis in the present scan.'"
    ),

    # --- WHOLE BODY PET-CT (Strict Verbatim Extraction) ---
    "{{pet_date}}": "Extract the Date of the PET-CT scan.",

    "{{pet_brain}}": (
        "Extract the 'Brain' section of the PET-CT **VERBATIM** (Word-for-Word). "
        "**STRICT RULE:** Do not summarize. Include every sentence regarding parenchyma, ventricles, and any signal intensity or uptake mentioned."
    ),

    "{{pet_head_neck}}": (
        "Extract the 'Head & Neck' section of the PET-CT **VERBATIM**. "
        "Include all text regarding lymph nodes, thyroid, and other structures. "
        "Capture every detail about 'avid' or 'non-avid' lesions exactly as written."
    ),

    "{{pet_chest}}": (
        "Extract the 'Chest' section of the PET-CT **VERBATIM**. "
        "**CRITICAL:** You must transcribe every single measurement (cm), SUV max value, and anatomical location exactly. "
        "Include all text descriptions of lesions, nodules, lymph nodes, lungs, heart, and pleural spaces. "
        "Do not omit any finding, even if it is described as 'questionable' or 'insignificant'."
    ),

    "{{pet_abdomen}}": (
        "Extract the 'Abdomen' section of the PET-CT **VERBATIM**. "
        "Include all findings for every organ listed (liver, spleen, pancreas, kidneys, etc.). "
        "Capture all mentions of calcifications, surgical sites, fluid, and physiological uptake exactly as they appear."
    ),

    "{{pet_bone}}": (
        "Extract the 'Bone' or 'Skeleton' section of the PET-CT **VERBATIM**. "
        "Include all details of lesions (lytic, sclerotic, etc.), specific vertebral levels, and all SUV values mentioned."
    ),

    "{{pet_impression}}": (
        "Extract the 'IMPRESSION' or 'CONCLUSION' section exactly as written. "
        "Capture the full summary text and any TNM Staging provided at the end."
    ),

    # --- ECHOCARDIOGRAPHY (Strict Data & Verbatim Findings) ---
    "{{echo_date}}": "STRICTLY extract the Date of the Echocardiography report exactly as written (e.g., 24/06/25).",

    # --- ECHO MEASUREMENTS (Values Only) ---
    "{{echo_ivs}}": (
        "STRICTLY extract the numeric value and unit for 'IVS-d' ONLY. "
        "Example: '13 mm'. Do not include the label."
    ),

    "{{echo_lvid_d}}": (
        "STRICTLY extract the numeric value and unit for 'LVID-d' ONLY. "
        "Example: '37 mm'."
    ),

    "{{echo_lvid_s}}": (
        "STRICTLY extract the numeric value and unit for 'LVID-s' ONLY. "
        "Example: '26 mm'."
    ),

    "{{echo_lvpw_d}}": (
        "STRICTLY extract the numeric value and unit for 'LVPW-d' ONLY. "
        "Example: '13 mm'."
    ),

    "{{echo_lvef}}": (
        "STRICTLY extract the percentage value for 'LVEF' ONLY. "
        "Example: '60%'."
    ),

    # --- ECHO NARRATIVE (Strict Line-by-Line) ---
    "{{echo_findings}}": (
        "STRICTLY extract the text findings below the measurements VERBATIM. "
        "**FORMATTING RULE:** You must maintain a strict LINE-BY-LINE format exactly as written in the document. "
        "Do not combine sentences into a paragraph. Do not summarize.\n"
        "**Example Output:**\n"
        "No LVRWMA/LVEF-60%\n"
        "Grade 1 diastolic function.\n"
        "All chambers normally.\n"
        "Trace AR\n"
        "Diagnosis - Normal LV function"
    ),

    # --- MAMMOGRAPHY & USG CORRELATION (Strict Verbatim Extraction) ---
    "{{mammo_date}}": "Extract the Date of the Mammography/USG report exactly as written (e.g., 09/06/25).",

    "{{mammo_indication}}": (
        "Extract the 'Indication' and 'On Examination (O/E)' sections **VERBATIM** (Word-for-Word). "
        "**STRICT RULE:** Transcribe the full clinical history and every physical exam finding (lump size, consistency, position) exactly as written in the notes."
    ),

    "{{mammo_findings_general}}": (
        "Extract the general 'MAMMOGRAPHY FINDINGS' paragraph **VERBATIM**. "
        "Include the exact text describing breast parenchyma and the ACR Breast Density type."
    ),

    "{{mammo_right_breast}}": (
        "Extract the 'RIGHT BREAST' section under MAMMOGRAPHY findings **VERBATIM**. "
        "**CRITICAL:** You must transcribe every measurement (mm/cm), clock position, margin description (e.g., 'spiculated'), "
        "calcification detail, skin thickening measurement, and BIRADS classification exactly as it appears."
    ),

    "{{mammo_right_axilla}}": (
        "Extract the 'RIGHT AXILLA' section under MAMMOGRAPHY findings **VERBATIM**. "
        "Report the status (e.g., 'Clear' or lymph node details) exactly as written."
    ),

    "{{mammo_left_breast}}": (
        "Extract the 'LEFT BREAST' section under MAMMOGRAPHY findings **VERBATIM**. "
        "Include exact descriptions of calcifications, tomosynthesis findings, and BIRADS scores."
    ),

    "{{mammo_left_axilla}}": (
        "Extract the 'LEFT AXILLA' section under MAMMOGRAPHY findings **VERBATIM**. "
        "Transcribe whatever is written regarding the left axilla."
    ),

    # --- USG SPECIFIC SECTIONS ---
    "{{usg_right_breast}}": (
        "Extract the 'RIGHT BREAST' section under USG FINDINGS **VERBATIM**. "
        "Include exact lesion echogenicity, measurements, Doppler vascularity, Elastography values (strain ratio/kPa), and BIRADS score. "
        "Do not summarize."
    ),

    "{{usg_right_axilla}}": (
        "Extract the 'RIGHT AXILLA' section under USG FINDINGS **VERBATIM**. "
        "Include all text regarding lymph node measurements, loss of fatty hilum, and infiltration status."
    ),

    "{{usg_right_supra}}": (
        "Extract the 'RIGHT SUPRA-/INFRA-CLAVICULAR REGION' section **VERBATIM**. "
        "If stated as 'Clear', write 'Clear'."
    ),

    "{{usg_left_breast}}": (
        "Extract the 'LEFT BREAST' section under USG FINDINGS **VERBATIM**. "
        "Include the exact description of parenchyma and any focal lesions found."
    ),

    "{{usg_left_axilla}}": (
        "Extract the 'LEFT AXILLA' section under USG FINDINGS **VERBATIM**. "
        "Include specific measurements and cortical thickness details exactly as written."
    ),

    "{{usg_left_supra}}": (
        "Extract the 'LEFT SUPRA-/INFRA-CLAVICULAR REGION' section **VERBATIM**."
    ),

    # --- CONCLUSION ---
    "{{mammo_impression}}": (
        "Extract the 'IMPRESSION' section **VERBATIM**. "
        "Include the final BIRADS assessment and the malignancy risk percentage exactly as written."
    ),

    "{{mammo_advice}}": (
        "Extract the 'ADVICE' section **VERBATIM**. "
        "Example: 'HPE (histopathological examination) correlation advised'."
    ),

    # --- HISTOPATHOLOGY REPORT (Ultra-Strict Verbatim) ---
    "{{hpe_date}}": "STRICTLY extract the Date of the Histopathology Report exactly as written (e.g., 09/06/25).",

    "{{hpe_specimen}}": (
        "STRICTLY extract the 'Specimen Sent' text VERBATIM (word-for-word) from the document. "
        "Do not alter or summarize the specimen description."
    ),

    "{{hpe_diagnosis}}": (
        "STRICTLY extract the 'Clinical Diagnosis' text VERBATIM from the document. "
        "Write exactly what is listed under Clinical Diagnosis."
    ),

    "{{hpe_gross}}": (
        "STRICTLY extract the entire 'Gross' examination section VERBATIM, preserving all measurements and embedding codes. "
        "Include details on tissue core lengths, color, and codes like 'A1-A2' exactly as they appear."
    ),

    "{{hpe_microscopy}}": (
        "STRICTLY extract the 'Microscopy' section VERBATIM, preserving the exact list format and scoring details. "
        "You must include the main description (e.g., 'Invasive breast carcinoma') followed immediately by the specific scores (Tubular differentiation, Nuclear pleomorphism, Mitosis, Score, Overall Grade) line-by-line as written."
    ),

    "{{hpe_advice}}": (
        "STRICTLY extract the 'Advice' section VERBATIM from the report. "
        "Write exactly what is requested (e.g., 'Immunohistochemistry for ER, PR...')."
    ),

    "{{hpe_ihc_all}}": (
        "STRICTLY extract the entire 'Immunohistochemistry markers studied' section VERBATIM. "
        "**FORMATTING RULE:** You must list every marker result on a new line. Do not combine them into a paragraph. "
        "Include the Marker Name, Status (Positive/Negative), and Score exactly as written.\n"
        "**Example Output:**\n"
        "ER: Negative (0/8)\n"
        "PR: Negative (0/8)\n"
        "Her2neu: Positive (3+)\n"
        "Ki-67: 15%"
    ),

    # --- FNAC REPORT (Strict Verbatim Extraction) ---
    "{{fnac_date}}": "STRICTLY extract the Date of the FNAC report exactly as written (e.g., 02/06/25).",

    "{{fnac_diagnosis}}": (
        "STRICTLY extract the 'Clinical diagnosis' from the FNAC report VERBATIM. "
        "Example: 'Right breast carcinoma'."
    ),

    "{{fnac_site}}": (
        "STRICTLY extract the 'Site of FNAC' section VERBATIM. "
        "**FORMATTING RULE:** You must preserve the exact numbering used in the document.\n"
        "Example:\n"
        "1. Right breast\n"
        "2. Right axilla"
    ),

    "{{fnac_microscopy}}": (
        "STRICTLY extract the 'Microscopy' section VERBATIM. "
        "**CRITICAL:** Preserve the numbering for each site extracted (1, 2, etc.) and transcribe the full text description for each.\n"
        "Do not summarize 'low cellularity' or 'malignant cells'—write the full sentence exactly as it appears."
    ),

    # --- USG ABDOMEN REPORT (Strict Verbatim Extraction) ---
    "{{usg_abdomen_date}}": "STRICTLY extract the Date of the USG Abdomen report exactly as written.",

    "{{usg_abdomen_findings}}": (
        "STRICTLY extract the entire 'ABDOMEN' findings section VERBATIM. "
        "**RULES:**\n"
        "1. Transcribe every sentence regarding the Liver, Gallbladder, CBD, Portal Vein, Kidneys, Spleen, Pancreas, and Bladder.\n"
        "2. Include all measurements (e.g., '15.5 cm') and observations (e.g., 'moderate fatty changes') exactly as written.\n"
        "3. Preserve the original bullet points or line breaks."
    ),

    "{{usg_abdomen_impression}}": (
        "STRICTLY extract the 'Impression' of the USG Abdomen VERBATIM. "
        "Example: 'Hepatomegaly with Moderate fatty changes in Liver'."
    ),

    # --- PROCEDURE & INTRAOP (Strict Formal) ---
    "{{procedure_details}}": (
        "State the exact procedure name, laterality, anesthesia, and date in one line. "
        "**FORMAT:** '[Procedure name], [side] under [anesthesia] on [DD/MM/YYYY].' "
        "Keep it crisp and standard. Do not add extra narrative unless explicitly provided."
    ),

    "{{intraop_findings}}": (
        "Rewrite the findings in formal operative-note style, using bullet points. "
        "**RULES:**\n"
        "1. **Primary Lesion:** Start with size, consistency, margins, quadrant/clock position, skin involvement (ulcer/peau d’orange), and fixation (skin/pectoralis/pectoral fascia).\n"
        "2. **Axilla:** Describe lymph node status with level and size (e.g., 'Level I nodes enlarged/hard, 1x1 cm').\n"
        "3. **Negatives:** Mention key negative findings explicitly if provided (e.g., 'not fixed to pectoral fascia').\n"
        "4. **Formatting:** Keep measurements in cm, preserve clock positions, and avoid informal words.\n"
        "5. **Extent:** End with the extent of surgery performed if stated (e.g., 'Level I and II ALND performed.').\n"
        "**CONSTRAINT:** Do not invent margins, vascularity, or bleeding if not stated."
    ),

    # --- HOSPITAL COURSE (Formal Medical Style) ---
    "{{hospital_course}}": (
        "Write a formal 'Course in Hospital' section. "
        "**RULES:**\n"
        "1. **Style:** Use past tense, objective medical language. Write as a single coherent paragraph.\n"
        "2. **Timeline:** Start with Presentation to OPD -> Evaluation/Diagnosis -> Prior Treatment (NACT) -> Surgery (Date & Anesthesia) -> Post-op Course.\n"
        "3. **Post-Op Details:** You MUST include daily progress if available (e.g., 'Drain output on POD 1 was 20 mL serosanguinous...'). If numbers conflict, choose the most consistent sequence.\n"
        "4. **Complications:** If complications occurred, detail them. If none, strictly state 'Post-operative course was uneventful'.\n"
        "5. **Discharge Statement:** If only the patient is discharged, end with: 'At present, patient is hemodynamically stable, ambulatory, with [Wound Status] and [Drain Status] and discharged on this date. (e.g.- 29/09/2026)'\n"
        "**ABBREVIATIONS:** Use standard terms like POD, GA, S/P, NACT, MRM, HTN appropriately."
    ),

    # --- DISCHARGE CONDITION ---
    "{{discharge_condition}}": (
        "Extract 'Condition at Discharge' verbatim from the notes. "
        "**RULE:** Use the exact medical language written (e.g., 'Patient is hemodynamically stable and ambulatory with healthy wound...'). "
        "**CRITICAL:** If the handwritten notes describe a different condition (e.g., 'Febrile', 'Wound dehiscence'), WRITE THAT DATA ONLY. "
        "If the section is empty, leave it empty."
    ),

    # --- MEDICATION & ADVICE ---
    "{{discharge_advice}}": (
        "Write specific medical instructions in professional language, Line-by-Line. "
        "CRITICAL: If the Hospital Course says 'discharged on X medication' or 'continue Y therapy', "
        "YOU MUST include that specific instruction here as a complete sentence.List the discharge medicines again with full dosage (e.g., Tab Pan 40mg - 1 tablet before breakfast)."
    ),

    "{{follow_up}}": "Plan & Follow-up details. Specify OPD Name, Days, and Tests to bring. YOU MUST include that specific instruction here as a complete sentence.",


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

    # --- NEW TABLE 1: CARDIAC & ACUTE PHASE ---
    # Anchor: {{CARDIAC_ANCHOR}}
    "{{cardiac_json}}": f"EXTRACT CARDIAC DATA as JSON. Keys: {', '.join(CARDIAC_TEST_ORDER)}.",

    # --- NEW TABLE 2: FLUID (CSF) ANALYSIS ---
    # Anchor: {{CSF_ANCHOR}}
    "{{csf_json}}": f"EXTRACT CSF DATA as JSON. Keys: {', '.join(CSF_TEST_ORDER)}.",

    
   "{{labs_json}}": (
        f"EXTRACT LABS as JSON for keys: {', '.join(SURGERY_TEST_ORDER)}. "
        "**CRITICAL COMBINATION RULE:** For combined rows, you MUST join values with a slash '/'. "
        "**EXAMPLES:**\n"
        "- 'tb_db': If Total Bilirubin is 0.34 and Direct is 0.08 -> output '0.34 / 0.08'\n"
        "- 'sgot_sgpt': If SGOT is 32 and SGPT is 37 -> output '32 / 37'\n"
        "- 'alp_ggt': If ALP is 129 and GGT is 40 -> output '129 / 40'\n"
        "- 'na_k': If Na is 141 and K is 4.1 -> output '141 / 4.1'\n"
        "- 'pt_inr': If PT is 14.5 and INR is 1.13 -> output '14.5 / 1.13'\n"
        "- 'thyroid': If T3 is 0.91, T4 is 11.2, TSH is 1.20 -> output '0.91 / 11.2 / 1.20'\n"
        "Format: {{'DD/MM': {{'hb': '11.1', 'na_k': '141/4.1'}}, ...}}"
    ),
}


# ==============================================================================
# SECTION D: THE LOGIC ENGINE
# ==============================================================================

def get_user_credentials():
    """Handles the Google Login popup and saves your token."""
    SCOPES = ['https://www.googleapis.com/auth/documents', 'https://www.googleapis.com/auth/drive']
    creds = None
    
    # --- FIX: Handle Empty/Corrupted token.json ---
    if os.path.exists('token.json'):
        try:
            creds = Credentials.from_authorized_user_file('token.json', SCOPES)
        except Exception:
            print("⚠️ token.json is corrupted/empty. Deleting it to re-login.")
            os.remove('token.json')
            creds = None
    # ----------------------------------------------
        
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
            except:
                # If refresh fails, force re-login
                creds = None
        
        if not creds:
            print("--- Launching Browser for Login... ---")
            flow = InstalledAppFlow.from_client_secrets_file(CLIENT_SECRET_FILE, SCOPES)
            creds = flow.run_local_server(port=0)
            
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
            
    return creds

def normalize_key(text):
    """Turns 'EsR', 'esr', 'ESR', 'Es_r' into just 'esr' for easier matching."""
    if not text: return ""
    return str(text).lower().replace("_", "").replace(" ", "").replace("-", "").strip()

def fill_smart_grid(service, doc_id, labs_data, test_order, anchor_name):
    """Fills grid with Robust Anchor Finding and Case-Insensitive Key Matching."""
    print(f"   -> Processing Grid for Anchor: {anchor_name}...")
    doc = service.documents().get(documentId=doc_id).execute(num_retries=5)
    content = doc.get('body').get('content')

    anchor_found = False
    table_index = -1
    anchor_row = -1
    anchor_col = -1
    
    # Remove brackets for the text search (e.g. "{{LAB_ANCHOR}}" -> "LAB_ANCHOR")
    clean_anchor_text = anchor_name.replace("{{", "").replace("}}", "").strip()

    # 1. Stitch Text Search (Your existing logic)
    for i, element in enumerate(content):
        if 'table' in element:
            for r_idx, row in enumerate(element['table']['tableRows']):
                for c_idx, cell in enumerate(row['tableCells']):
                    full_cell_text = ""
                    if 'content' in cell:
                        for content_item in cell['content']:
                            if 'paragraph' in content_item:
                                for element in content_item['paragraph']['elements']:
                                    if 'textRun' in element:
                                        full_cell_text += element['textRun']['content']
                    
                    # Check if the stitched text contains our specific anchor
                    if clean_anchor_text in full_cell_text:
                        anchor_found = True
                        table_index = i
                        anchor_row = r_idx
                        anchor_col = c_idx
                        print(f"      FOUND {anchor_name} at Row {r_idx}, Col {c_idx}")
                        break
                if anchor_found: break
        if anchor_found: break
    
    if not anchor_found:
        print(f"      WARNING: {anchor_name} NOT FOUND.")
        return []

    requests = []
    # 2. Clean the Anchor Text (Replace {{ANCHOR}} with empty space)
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
        target_col = anchor_col + i
        if target_col >= len(table['tableRows'][0]['tableCells']): break

        try:
            # A. Fill Date Header
            cell = table['tableRows'][anchor_row]['tableCells'][target_col]
            end_index = cell['endIndex'] - 1
            requests.append({'insertText': {'location': {'index': end_index}, 'text': str(date)}})
            
            # B. Fill Test Values (NEW: Case-Insensitive Logic)
            
            # Create a "normalized map" of the AI's data for this day
            # Example: AI gave {"EsR": "45"} -> we convert key to "esr": "45"
            day_map = {normalize_key(k): v for k, v in labs_data[date].items()}
            
            # Loop through YOUR specific test order list
            for test_idx, python_key in enumerate(test_order):
                target_row = anchor_row + 1 + test_idx
                if target_row >= len(table['tableRows']): break
                
                # Normalize the key we are looking for (e.g., 'HS Trop I' -> 'hstropi')
                search_key = normalize_key(python_key)
                
                # Retrieve value using the normalized key
                val = str(day_map.get(search_key, ""))
                
                # --- BRUTE FORCE CLEANING (Your existing logic) ---
                val = val.replace("{", "").replace("}", "").replace("[", "").replace("]", "").strip()

                if val and val != "":
                    try:
                        cell = table['tableRows'][target_row]['tableCells'][target_col]
                        end_index = cell['endIndex'] - 1
                        requests.append({'insertText': {'location': {'index': end_index}, 'text': val}})
                    except: pass
        except: continue

    # 5. Sort requests to prevent index shifting
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

    ### 0. GENERAL LAB GRID (Haemoglobin, Urea, etc):
    - Key: "{{labs_json}}". Keys: {SURGERY_TEST_ORDER}

    ### CRITICAL INSTRUCTION: SEMANTIC SEARCH
    Understand the **biological meaning** of tests. 
    - Map "CK-MB" or "Creatine Kinase-MB" to "cpkmb".
    - Map "GeneXpert" or "MTB/RIF" to "csfcbnaat".
    - Map "Cell Count" in CSF to "csftlc".


    ### 0. LOGIC FOR LAB DATA (NEW):
    - Extract ALL lab values associated with dates.
    - Return a JSON object under key "{{labs_json}}".
    - Keys must match EXACTLY: {SURGERY_TEST_ORDER}
    - Composite fields: 'dlc_diff' (N/L/M/E/B), 'indices' (MCV/MCH/MCHC), 'elyte' (Na/K).
    - If a value is missing for a specific date, DO NOT include that key.

    ### 1. LOGIC FOR "+ve / -ve" FIELDS:
    For placeholders {{pallor}}, {{icterus}}, {{cyanosis}}, {{clubbing}}, {{lymphadenopathy}}, {{edema}}:
    - READ the text carefully.
    - If the note says "Present", "Positive", "++" -> Output: "Present"
    - If the note says "Absent", "Negative", "--", "Nil", or DOES NOT MENTION it -> Output: "Absent"

    ### 2. LOGIC FOR SYSTEMIC EXAM (SMART MERGE):
    For CVS, RS, PA, and CNS, you have a "Normal Default" sentence.
    **Your Goal:** Start with the Default sentence. If the patient's notes mention a specific abnormal finding, REPLACE only that specific part of the sentence.

    * **CVS (Default: "S1 S2 heard, no murmurs"):**
        * If note says "Murmur present", Output: "S1 S2 heard, Murmur present".
        * If note says "Muffled sounds", Output: "Muffled heart sounds".

    * **Respiratory (Default: "B/L normal vesicular breath sounds heard"):**
        * If note says "Crepts in right base", Output: "Air entry present, Crepitations in right base".
        * If note says "Wheeze present", Output: "B/L Air entry present with Wheeze".
        * Change ANY part that is different in the source text.


    ### 3. RULES FOR Height and Weight:
    - Extract height number only , dont give units. 
    - Extract weight number only , dont give units .

    ### 3. RULES FOR RESIDENTS/FACULTY:
    - Extract exact titles like "(SR): DR NAME", "(JR): DR NAME". 
    - Keep them on separate lines.

    ### 4. RULES FOR VITALS (Admission vs Discharge):
    - **Admission Vitals:** Look for "Vitals at Admission" or early dates.
    - **Units:** Extract ONLY the number for Pulse, RR, Temp, SpO2.

   

    **Discharge Advice:**
       - Must be **Line-by-Line** using complete, professional sentences.
       - Write specific medical instructions in professional language, Line-by-Line.
       - CRITICAL: If the Hospital Course says 'discharged on X medication' or 'continue Y therapy
       - YOU MUST include that specific instruction here as a complete sentence.List the discharge medicines again with full dosage (e.g., Tab Pan 40mg - 1 tablet before breakfast).
       - Do not simply write "Review in OPD". Give specific care instructions relevant to the diagnosis.

    



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

                # --- ROBUST FIX START ---
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
        'mimeType': 'application/vnd.google-apps.document'  # Auto-convert to Doc
    }
    
    try:
        copy_response = drive_service.files().copy(
            fileId=MASTER_TEMPLATE_ID,
            body=file_metadata,
            supportsAllDrives=True
        ).execute(num_retries=5)
        NEW_DOCUMENT_ID = copy_response.get('id')
        print(f"--> File Created! ID: {NEW_DOCUMENT_ID}")

    except Exception as e:
        print(f"Drive Error: {e}")
        # Return the error to the frontend so we can see it!
        return {"error": f"Google Drive Permission Error: {str(e)}"}

    print(f"--- 4. Filling Data... ---")
    docs_service = build('docs', 'v1', credentials=creds)

    # --- 1. Fill Main Lab Grid (SAFE MODE) ---
    if "{{labs_json}}" in extracted_data:
        lab_data = extracted_data["{{labs_json}}"]
        
        # Check if it is a dictionary (using your existing safe logic)
        if isinstance(lab_data, dict):
            # <--- CHANGE IS HERE: Use SURGERY_TEST_ORDER instead of LAB_TEST_ORDER
            lab_requests = fill_smart_grid(
                docs_service, 
                NEW_DOCUMENT_ID, 
                lab_data, 
                SURGERY_TEST_ORDER,  # <--- CRITICAL CHANGE
                "{{LAB_ANCHOR}}"
            )
            
            if lab_requests:
                docs_service.documents().batchUpdate(
                    documentId=NEW_DOCUMENT_ID, 
                    body={'requests': lab_requests}
                ).execute(num_retries=5)
                print("   -> Surgery Lab Grid Filled.")
        else:
            print(f"   ⚠️ Skipping Labs: AI returned {type(lab_data)}")
        
        # Cleanup
        del extracted_data["{{labs_json}}"]

    # --- 2. Fill Cardiac Grid (SAFE MODE) ---
    if "{{cardiac_json}}" in extracted_data:
        cardiac_data = extracted_data["{{cardiac_json}}"]
        if isinstance(cardiac_data, dict):
            cardiac_requests = fill_smart_grid(docs_service, NEW_DOCUMENT_ID, cardiac_data, CARDIAC_TEST_ORDER, "{{CARDIAC_ANCHOR}}")
            if cardiac_requests:
                docs_service.documents().batchUpdate(documentId=NEW_DOCUMENT_ID, body={'requests': cardiac_requests}).execute(num_retries=5)
                print("   -> Cardiac Grid Filled.")
        else:
            print(f"   ⚠️ Skipping Cardiac: AI returned {type(cardiac_data)} instead of Dict")
        del extracted_data["{{cardiac_json}}"]

    # --- 3. Fill CSF Grid (SAFE MODE) ---
    if "{{csf_json}}" in extracted_data:
        csf_data = extracted_data["{{csf_json}}"]
        if isinstance(csf_data, dict):
            csf_requests = fill_smart_grid(docs_service, NEW_DOCUMENT_ID, csf_data, CSF_TEST_ORDER, "{{CSF_ANCHOR}}")
            if csf_requests:
                docs_service.documents().batchUpdate(documentId=NEW_DOCUMENT_ID, body={'requests': csf_requests}).execute(num_retries=5)
                print("   -> CSF Grid Filled.")
        else:
            print(f"   ⚠️ Skipping CSF: AI returned {type(csf_data)} instead of Dict")
        del extracted_data["{{csf_json}}"]

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
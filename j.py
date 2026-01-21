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
    """
    Extracts usage tokens, calculates cost, and saves to a CSV file.
    """
    try:
        # 1. Get Token Counts directly from the API response
        # Google sends this in 'usage_metadata'
        usage = response.usage_metadata
        in_tokens = usage.prompt_token_count
        out_tokens = usage.candidates_token_count
        total_tokens = usage.total_token_count
        
        # 2. Calculate Cost
        rates = PRICING.get(model_name, {"input": 0, "output": 0})
        cost_input = (in_tokens / 1_000_000) * rates["input"]
        cost_output = (out_tokens / 1_000_000) * rates["output"]
        total_cost = cost_input + cost_output
        
        # 3. Prepare the Data Row
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        row = [
            timestamp, 
            model_name, 
            note,
            in_tokens, 
            out_tokens, 
            total_tokens, 
            f"${total_cost:.5f}",
            f"₹{total_cost * 87:.2f}" # Convert to INR
        ]
        
        # 4. Save to CSV
        file_path = "cost_history.csv"
        file_exists = Path(file_path).exists()
        
        with open(file_path, "a", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            # Write Header if new file
            if not file_exists:
                writer.writerow(["Time", "Model", "Task", "Input Tokens", "Output Tokens", "Total Tokens", "Cost (USD)", "Cost (INR)"])
            
            writer.writerow(row)
            
        print(f"   [💰 Cost Logged: ₹{total_cost * 87:.2f} for {note}]")
        
    except Exception as e:
        print(f"   [⚠️ Cost Logging Failed: {e}]")

# ==============================================================================
# SECTION A: CONFIGURATION (Loaded from .env file)
# ==============================================================================
load_dotenv()

# 2. Get keys safely
GENAI_API_KEY = os.getenv("GENAI_API_KEY")
MASTER_TEMPLATE_ID = os.getenv("MASTER_TEMPLATE_ID")
OUTPUT_FOLDER_ID = os.getenv("OUTPUT_FOLDER_ID")
CLIENT_SECRET_FILE = os.getenv("CLIENT_SECRET_FILE")

# 3. Safety Check
if not GENAI_API_KEY or not MASTER_TEMPLATE_ID:
    print("CRITICAL ERROR: Keys are missing.")
    print("Make sure you created the '.env' file with your API keys.")
    exit()

# ==============================================================================


# ==============================================================================
# SECTION B.2: LAB TEST ORDER (Must match your Table Rows exactly)
# ==============================================================================
LAB_TEST_ORDER = [
    "hb", "tlc", "dlc_diff", "indices", "plt", "rdw", "b_total", "b_direct", 
    "sgpt", "sgot", "alp", "ggt", "protein", "albumin", "globulin", 
    "urea", "cr", "na", "k", "cl", "ca", "uric_acid", 
    "pt_inr", "aptt", "procal", "crph"
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

    # ---  ADMISSION DETAILS ---
    "{{doa}}": "Date of Admission (DOA)",
    "{{dod}}": "Date of Discharge (DOD)",
    "{{residents}}": "List of Residents (SR/JR). Maintain formatting (e.g., (SR): Name).",
    "{{faculty}}": "List of Treating Faculty.",
    
    # --- DIAGNOSIS & SUMMARY ---
    "{{final_diagnosis}}": "Extract Final Diagnosis completely. Maintain medical capitalization.List Diagnosis serially (1., 2., 3.). DO NOT use + signs. **FORBIDDEN CHARACTER:** Do NOT use the + sign. BAD: Dengue +ve, T2DM + HTN. GOOD:** Dengue Positive, Type 2 Diabetes Mellitus and Hypertension. Must be a Numbered List (1., 2., 3.).",
    "{{case_summary}}": "Extract Case Summary. Include past history, presenting complaints, and duration.""Write a PROFESSIONAL NARRATIVE. "
        "Start with: 'Name is a Age-year-old Gender, resident of Address...' "
        "Include comorbidities, family history, and presenting complaints numbered serially.",
    
    # --- EXAM (GENERAL) - THE +/- VE LOGIC ---
    "{{pallor}}": "Look for Pallor. If present -> '+ve'. If absent/normal -> '-ve'.",
    "{{icterus}}": "Look for Icterus. If present -> '+ve'. If absent/normal -> '-ve'.",
    "{{cyanosis}}": "Look for Cyanosis. If present -> '+ve'. If absent/normal -> '-ve'.",
    "{{clubbing}}": "Look for Clubbing. If present -> '+ve'. If absent/normal -> '-ve'.",
    "{{lymphadenopathy}}": "Look for Lymphadenopathy. If present -> '+ve'. If absent/normal -> '-ve'.",
    "{{edema}}": "Look for Pedal Edema. If present -> '+ve'. If absent/normal -> '-ve'.",

    # --- VITALS ---
    "{{pulse}}": "Extract Pulse Rate number only (e.g. 101).",
    "{{rr}}": "Extract Respiratory Rate (RR) number only.",
    "{{bp}}": "Extract Blood Pressure (BP) number only",
    "{{temp}}": "Extract Temperature, number only",
    "{{spo2}}": "Extract SpO2 percentage, number only",

    # --- SYSTEMIC EXAM (SMART DEFAULTS) ---
    "{{cvs_exam}}": "CVS Findings. If normal, use default 'S1, S2 +'. If abnormal, describe findings.",
    "{{rs_exam}}": "Respiratory Findings. If normal, use default 'B/L Air entry present'.",
    "{{pa_exam}}": "Per Abdomen Findings. If normal, use default 'Non distended, non-tender, bowel sounds are present, and no palpable organomegaly Present.'.",
    "{{cns_exam}}": "CNS Findings. If normal, use default 'EV3M5, Pupils: B/L NSNR,,meningeal signs absent. Motor power cannot be assessed.'.",

    
    # --- COURSE (UPDATED FOR DETAILED NARRATIVE) ---
    "{{hospital_course}}": (
        "Write a COMPREHENSIVE MEDICAL NARRATIVE. "
        "Start with admission complaints and initial assessment (GCS, Vitals). "
        "Detail the diagnostic journey: Mention POSITIVE findings from Imaging/Labs and pertinent NEGATIVES. "
        "Describe the management: Antibiotics, Procedures (e.g., Dialysis, CVC), and treatment of complications (e.g., DVT, AKI). "
        "Include significant abnormal lab values to show trends (e.g., 'Creatinine peaked at 2.4'). "
        "Conclude with the patient's clinical status at discharge."
    ),

    # ---  DISCHARGE VITALS  ---
    "{{dis_gcs}}": "Discharge GCS/Sensorium (Look for E4V5M6 pattern). If missing, leave blank.",
    "{{dis_pulse}}": "Discharge Pulse (Number only).If missing, leave blank.",
    "{{dis_bp}}": "Discharge BP (Number only).If missing, leave blank.",
    "{{dis_rr}}": "Discharge RR (Number only).If missing, leave blank.",
    "{{dis_temp}}": "Discharge Temp (Number only).If missing, leave blank.",
    "{{dis_spo2}}": "Discharge SpO2 (Number only).If missing, leave blank.",

    # --- MEDICATION & ADVICE ---
    "{{treatment_given}}": "You MUST list every medicine name, dosage, and frequency explicitly line-by-line. KEEP LINE-BY-LINE FORMAT. Also Include Injections and IV fluids",
    "{{discharge_advice}}": (
        "Write specific medical instructions in professional language, Line-by-Line. "
        "CRITICAL: If the Hospital Course says 'discharged on X medication' or 'continue Y therapy', "
        "YOU MUST include that specific instruction here as a complete sentence.List the discharge medicines again with full dosage (e.g., Tab Pan 40mg - 1 tablet before breakfast)."
    ),
    "{{general_advice}}": "General Advice (Diet/Warning signs). General Advice in bullet points. "
        "Include: 'TAKE MEDICINES REGULARLY', 'WATCH FOR SIDE EFFECTS', 'DIET INSTRUCTIONS'.",
    "{{follow_up}}": "Plan & Follow-up details. Specify OPD Name, Days, and Tests to bring. YOU MUST include that specific instruction here as a complete sentence.",

    # --- 8. SEROLOGY ---
    "{{hbsag}}": "HBsAg result (e.g. Negative or Positive only).",
    "{{hiv}}": "Anti HIV I & II result(e.g. Negative or Positive only).",
    "{{hcv}}": "Anti HCV result(e.g. Negative or Positive only).",

    # --- 9. URINE RM ---
    "{{urine_date}}": "Date of Urine RM test (DD/MM/YY).",
    "{{urine_pus}}": "Urine Pus cells value.",
    "{{urine_epi}}": "Urine Epithelial cells value.",
    "{{urine_rbc}}": "Urine RBC value.",
    "{{urine_casts}}": "Urine Casts value.",

    # --- 10. SPECIAL LABS (TOXO / BAL) ---
    "{{toxo_date}}": "Date of TOXO IgM/Cryptococcus LFA.",
    "{{toxo_igm}}": "TOXO IgM result(e.g. Negative or Positive only).",
    "{{crypto_lfa}}": "Cryptococcus LFA result(e.g. Negative or Positive only).",
    
    "{{bal_date}}": "Date of BAL tests.",
    "{{bal_fungal}}": "BAL Fungal Culture result.",
    "{{bal_gram}}": "BAL Gram Stain result.",

    # --- 11. METABOLIC & WORKUP ---
    "{{sugar_f_pp}}": "Fasting/Post Prandial Glucose value.",
    "{{hba1c}}": "HbA1c value.",
    "{{tsh}}": "TSH value.",
    "{{vit_d}}": "Vitamin D value.",
    "{{ipth}}": "iPTH value.",
    "{{tc}}": "Total Cholesterol (TC) value.",
    "{{tg}}": "Triglycerides (TG) value.",
    "{{hdl}}": "HDL value.",
    "{{ldl}}": "LDL value.",
    "{{ps_rbc}}": "Peripheral Smear RBC description.",
    "{{ps_wbc}}": "Peripheral Smear WBC description.",
    "{{ps_plt}}": "Peripheral Smear Platelet description.",
    "{{retic}}": "Reticulocyte Count / RPI.",
    "{{workup_indices}}": "MCV/MCH/MCHC (from Workup section).",
    "{{iron}}": "Iron value.",
    "{{tsat}}": "TSAT value.",
    "{{tibc}}": "TIBC value.",
    "{{ferritin}}": "Ferritin value.",
    "{{vit_b12}}": "Vitamin B12 value.",
    "{{folate}}": "Folate value.",
    "{{ldh}}": "LDH value.",
    "{{coombs}}": "DAT/IAT result.",
    "{{stool_obt}}": "Stool OBT result.",

   
    # --- NEW TABLE 1: CARDIAC & ACUTE PHASE ---
    # Anchor: {{CARDIAC_ANCHOR}}
    "{{cardiac_json}}": f"EXTRACT CARDIAC DATA as JSON. Keys: {', '.join(CARDIAC_TEST_ORDER)}.",

    # --- NEW TABLE 2: FLUID (CSF) ANALYSIS ---
    # Anchor: {{CSF_ANCHOR}}
    "{{csf_json}}": f"EXTRACT CSF DATA as JSON. Keys: {', '.join(CSF_TEST_ORDER)}.",

    # --- NEW TABLE 3: CULTURES (Specific Cells) ---
    "{{blood_cs_date}}": "Date of Blood Culture (Sent on).",
    "{{blood_cs_res}}": "Result of Blood Culture (Growth/Organism).",
    "{{urine_cs_date}}": "Date of Urine Culture (Sent on).",
    "{{urine_cs_res}}": "Result of Urine Culture (Growth/Organism).",

    # --- IMAGING PLACEHOLDERS (Add to Section C) ---
    "{{date_ncct}}": "Date of NCCT Brain or CT Head",
    "{{ncct_findings}}": "NCCT Brain Findings (Verbatim). Do not summarize.",
    "{{ncct_imp}}": "NCCT Impression",
    
    "{{date_mri}}": "Date of MRI Brain",
    "{{mri_findings}}": "MRI Brain Findings (Verbatim). Do not summarize.",
    "{{mri_imp}}": "MRI Impression",
    
    "{{date_bronch}}": "Date of Bronchoscopy",
    "{{bronch_findings}}": "Bronchoscopy Findings (Verbatim). Do not summarize.",
    "{{bronch_imp}}": "Bronchoscopy Impression",
    
    "{{date_doppler}}": "Date of USG Doppler",
    "{{doppler_findings}}": "Doppler Findings (Verbatim). Do not summarize.",
    "{{doppler_imp}}": "Doppler Impression",
    
    "{{date_cect}}": "Date of CECT Thorax/Abdomen",
    "{{thorax_findings}}": "CECT Thorax Findings (Verbatim). Do not summarize.",
    "{{abdomen_findings}}": "CECT Abdomen Findings (Verbatim). Do not summarize.",
    "{{cect_imp}}": "CECT Impression",

   # --- LAB DATA EXTRACTION ---
    "{{labs_json}}": f"EXTRACT LABS as JSON. Keys: {', '.join(LAB_TEST_ORDER)}. Format: {{'DD/MM': {{'hb': '10', 'urea': '40'}}, ...}}"
}


# ==============================================================================
# SECTION C.2: PROFESSIONAL IMAGING DEFAULTS
# ==============================================================================
IMAGING_DEFAULTS = {
    # --- 1. NCCT BRAIN ---
    "{{date_ncct}}": "", 
    "{{ncct_findings}}": (
        "Parenchyma: Both cerebral hemispheres show normal attenuation values. Grey-white matter differentiation is well preserved. No focal or diffuse areas of altered density are seen.\n"
        "Ventricular System: The supratentorial and infratentorial ventricular systems are normal in size, shape, and position. There is no midline shift or mass effect.\n"
        "Deep Structures: The basal ganglia, thalami, and internal capsule appear normal bilaterally.\n"
        "Posterior Fossa: The cerebellum and brainstem appear normal.\n"
        "Bone & Sinuses: The bony calvarium is intact. Visualised paranasal sinuses and mastoid air cells are clear and aerated."
    ),
    "{{ncct_imp}}": "Normal NCCT Brain study. No significant intracranial abnormality detected.",

    # --- 2. CEMRI BRAIN ---
    "{{date_mri}}": "",
    "{{mri_findings}}": (
        "Signal Intensity: Normal signal intensity is noted in the brain parenchyma on all sequences. No areas of restricted diffusion or abnormal susceptibility blooming are seen.\n"
        "Contrast Enhancement: Post-contrast administration, there is no evidence of abnormal parenchymal enhancement, ring-enhancing lesions, or leptomeningeal enhancement.\n"
        "Ventricles: The ventricular system is normal in caliber. No periventricular ooze or ependymal enhancement is noted.\n"
        "Sella/Pituitary: The pituitary gland and sella turcica appear normal.\n"
        "Vessels: Major intracranial flow voids are preserved.\n"
        "Extracranial: Visualised orbits, sinuses, and calvarium are unremarkable."
    ),
    "{{mri_imp}}": "Normal Contrast-Enhanced MRI Brain. No evidence of meningitis, granuloma, or acute infarct.",

    # --- 3. BRONCHOSCOPY ---
    "{{date_bronch}}": "",
    "{{bronch_findings}}": (
        "Upper Airway: Upper respiratory tract anatomy is normal.\n"
        "Vocal Cords: Normal appearance; bilateral cords are equal and mobile. No palsy or growth seen.\n"
        "Trachea: Normal caliber and mucosa. No secretions or narrowing.\n"
        "Carina: Main carina is sharp, central, and normal.\n"
        "Right Bronchial Tree: All segmental bronchi visualized. Mucosa is normal. No secretions, nodularity, or endobronchial mass lesions noted.\n"
        "Left Bronchial Tree: All segmental bronchi visualized. Mucosa is normal. No secretions, nodularity, or endobronchial mass lesions noted."
    ),
    "{{bronch_imp}}": "Normal Bronchoscopy study.",

    # --- 4. USG DOPPLER ---
    "{{date_doppler}}": "",
    "{{doppler_findings}}": (
        "Vessel Architecture: The bilateral Common Femoral Vein (CFV), Superficial Femoral Vein (SFV), and Popliteal Veins are well visualized with normal caliber.\n"
        "Compressibility: Complete compressibility is seen in all visualized venous segments (indicating patency).\n"
        "Flow Dynamics: Color Doppler demonstrates normal spontaneous phasic flow with respiration. Good flow augmentation is seen on distal compression.\n"
        "Lumen: No evidence of echogenic thrombus or filling defects within the lumen.\n"
        "Soft Tissue: Visualised subcutaneous planes appear normal with no edema."
    ),
    "{{doppler_imp}}": "No sonographic evidence of Deep Vein Thrombosis (DVT) in the bilateral lower limbs.",

    # --- 5. CECT THORAX & ABDOMEN ---
    "{{date_cect}}": "",
    "{{thorax_findings}}": (
        "Lungs: Lung parenchyma is clear bilaterally. No nodules, consolidation, cavitation, or tree-in-bud opacities seen.\n"
        "Pleura: No evidence of pleural effusion or pneumothorax.\n"
        "Mediastinum: Central trachea. No significant mediastinal or hilar lymphadenopathy.\n"
        "Cardiovascular: Heart size is within normal limits. Major thoracic vessels are normal."
    ),
    "{{abdomen_findings}}": (
        "Liver: Normal size (approx. 12-14 cm), shape, and homogeneous attenuation. No focal lesions. No intrahepatic biliary radical dilatation.\n"
        "Gall Bladder: Normal distension and wall thickness. No radio-opaque calculi. CBD is normal.\n"
        "Pancreas & Spleen: Normal size and texture. No focal lesions, calcifications, or ductal dilatation.\n"
        "Kidneys: Bilateral kidneys are normal in size and position. Corticomedullary differentiation is preserved. No hydronephrosis or calculi.\n"
        "Bowel: Visualized bowel loops show normal wall thickness. No signs of obstruction or mass lesions.\n"
        "Lymph Nodes: No significant retroperitoneal or mesenteric lymphadenopathy.\n"
        "Peritoneum: No free fluid (ascites) seen in the abdomen or pelvis."
    ),
    "{{cect_imp}}": "Normal CECT Thorax and Abdomen study. No evidence of infective etiology or malignancy."
}

# ==============================================================================
# SECTION D: THE LOGIC ENGINE
# ==============================================================================

def get_user_credentials():
    """Handles the Google Login popup and saves your token."""
    SCOPES = ['https://www.googleapis.com/auth/documents', 'https://www.googleapis.com/auth/drive']
    creds = None
    
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
        
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
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
    - Key: "{{labs_json}}". Keys: {LAB_TEST_ORDER}

    ### CRITICAL INSTRUCTION: SEMANTIC SEARCH
    Understand the **biological meaning** of tests. 
    - Map "CK-MB" or "Creatine Kinase-MB" to "cpkmb".
    - Map "GeneXpert" or "MTB/RIF" to "csfcbnaat".
    - Map "Cell Count" in CSF to "csftlc".
    
  2. **CARDIAC & INFLAMMATORY** ("{{cardiac_json}}"):
       - **Search Logic:** Map "Trop I" -> "hstropi", "CK-MB" -> "cpkmb", "Total CK" -> "cpknac", "NT-proBNP" -> "bnp", "PCT" -> "procal".
       - Target Keys: {CARDIAC_TEST_ORDER}
       - Format: {{ "Date": {{ "hstropi": "0.01", "esr": "120" }} }}

    3. **CSF / FLUID ANALYSIS** ("{{csf_json}}"):
       - **Target Keys:** {CSF_TEST_ORDER}
       - **Search Logic:**
         * "tlc": Look for "Cell Count", "Total Cells", "WBCs".
         * "dlc": Look for "Differential", "Polymorphs/Lymphocytes".
         * "cbnaat": Look for "GeneXpert", "Xpert MTB/RIF".
         * "glucose": Look for "Sugar", "Glu".
         * "culture": Look for "Aerobic Culture", "Pyogenic Culture".
       - *Format:* {{ "25/12": {{ "tlc": "5 cells", "glucose": "45" }} }}

    ### 1. SPECIAL TESTS:
    - **Cultures:** Extract Date & Result. Look for "Blood C/S", "Urine Culture".
    - **Metabolic:** Look for "Lipid Profile" (TC/TG/HDL/LDL).
    ### 3. CULTURE REPORTS (Specific Placeholders):
    - **Blood C/S:** Extract "Sent on" Date -> {{blood_cs_date}}. Extract Result -> {{blood_cs_res}}.
    - **Urine C/S:** Extract "Sent on" Date -> {{urine_cs_date}}. Extract Result -> {{urine_cs_res}}.
    - If "Skin commensal" or "Contaminant", write that.
    - If "No growth", write "No growth".

    ### 0.1 SPECIAL TEST DATES & WORKUP:
    - **Urine/BAL/Toxo Dates:** Extract the date written next to or above these specific test blocks. 
    - **Workup Values:** Extract the value *only*. If a test (like 'Stool OBT') is not mentioned or has no value, leave it empty.
    - **Format:** For Urine/BAL/Toxo dates, return ONLY the date (e.g., "27/12/25"). Do NOT add brackets in the output (the template has them).

    ### 0. LOGIC FOR LAB DATA (NEW):
    - Extract ALL lab values associated with dates.
    - Return a JSON object under key "{{labs_json}}".
    - Keys must match EXACTLY: {LAB_TEST_ORDER}
    - Composite fields: 'dlc_diff' (N/L/M/E/B), 'indices' (MCV/MCH/MCHC), 'elyte' (Na/K).
    - If a value is missing for a specific date, DO NOT include that key.

    ### 1. LOGIC FOR "+ve / -ve" FIELDS:
    For placeholders {{pallor}}, {{icterus}}, {{cyanosis}}, {{clubbing}}, {{lymphadenopathy}}, {{edema}}:
    - READ the text carefully.
    - If the note says "Present", "Positive", "++" -> Output: "+ve"
    - If the note says "Absent", "Negative", "--", "Nil", or DOES NOT MENTION it -> Output: "-ve"

    ### 2. LOGIC FOR SYSTEMIC EXAM (SMART MERGE):
    For CVS, RS, PA, and CNS, you have a "Normal Default" sentence.
    **Your Goal:** Start with the Default sentence. If the patient's notes mention a specific abnormal finding, REPLACE only that specific part of the sentence.

    * **CVS (Default: "S1, S2 +"):**
        * If note says "Murmur present", Output: "S1, S2 +, Murmur present".
        * If note says "Muffled sounds", Output: "Muffled heart sounds".

    * **Respiratory (Default: "B/L Air entry present"):**
        * If note says "Crepts in right base", Output: "Air entry present, Crepitations in right base".
        * If note says "Wheeze present", Output: "B/L Air entry present with Wheeze".
        * Change ANY part that is different in the source text.

    * **Per Abdomen (Default: "Non distended, non-tender, bowel sounds are present, and no palpable organomegaly Present."):**
        * If note says "Distended", Output: "Distended, non-tender, bowel sounds are present...".
        * If note says "Hepatosplenomegaly", Output: "Non-distended, non-tender,bowel sounds are present, Hepatosplenomegaly present".
        * If note says "Absent bowel sounds", Output: "Non-distended, non-tender, Bowel sounds absent and no palpable organomegaly Present".
        * *Rule:* Treat "Soft", "Tenderness", "Bowel Sounds", and "Organomegaly" as 4 separate switches. Change only what is mentioned.

    * **CNS (Component-Based Smart Merge):**
        * **The Base Sentence:** "Conscious, Oriented (E4V5M6). Pupils B/L NSNR. Motor Power 5/5 in all limbs. No focal deficit."
        * **Instruction:** Treat this as 4 separate components. Update ONLY the component that is mentioned in the text.
        * **Component 1: Consciousness (Default: Conscious/E4V5M6)** - If text says "Drowsy" or "E3V4M5" -> Change ONLY this part.
        * **Component 2: Pupils (Default: B/L NSNR)** - If text says "Right pupil dilated" -> Change ONLY this part.
        * **Component 3: Motor Power (Default: 5/5)** - If text says "Left hemiparesis" or "Power 3/5" -> Change ONLY this part.
        * *Result:* A drowsy patient with good power will read: "Drowsy (E3V4M5). Pupils B/L NSNR. Motor Power 5/5 in all limbs. No focal deficit."

    ### 3. RULES FOR RESIDENTS/FACULTY:
    - Extract exact titles like "(SR): DR NAME", "(JR): DR NAME". 
    - Keep them on separate lines.

    ### 4. RULES FOR VITALS (Admission vs Discharge):
    - **Admission Vitals:** Look for "Vitals at Admission" or early dates.
    - **Discharge Vitals:** Look specifically for "Vitals at discharge" or "On discharge".
    - **GCS:** Look for patterns like "E4V5M6" or "E4 V5 M6" under discharge vitals.
    - **Units:** Extract ONLY the number for Pulse, RR, Temp, SpO2 (e.g., extract "98", not "98 F").

    ### 5. RULES FOR LISTS (Treatment/Advice):
    - **CRITICAL:** Do NOT combine medications into a paragraph.
    - Keep them strictly **Line-by-Line** (one medicine per line).
    - Maintain the dosage and frequency (e.g. "TAB PAN 40 MG OD").

    **Discharge Advice:**
       - Must be **Line-by-Line** using complete, professional sentences.
       - Write specific medical instructions in professional language, Line-by-Line.
       - CRITICAL: If the Hospital Course says 'discharged on X medication' or 'continue Y therapy
       - YOU MUST include that specific instruction here as a complete sentence.List the discharge medicines again with full dosage (e.g., Tab Pan 40mg - 1 tablet before breakfast).
       - Do not simply write "Review in OPD". Give specific care instructions relevant to the diagnosis.

    **Hospital Course (SMART NARRATIVE):**
       - **Synthesize** the course from the notes. Do not just copy bullet points.
       - **Focus on Findings:** Highlight what was FOUND (Positives) and what was ruled out (Negatives).
       - **Key Data Only:** Mention important abnormal values (e.g., "Troponin was elevated at 0.85"), but do NOT list the date of every single routine test unless it marks a turning point.
       - **Flow:** Admission -> Workup -> Treatment -> Complications -> Recovery.
       - **Language:** Use formal phrases like "Patient was initiated on...", "Course was complicated by...", "Evaluation revealed...".

    ### 7. IMAGING REPORTS (Verbatim Extraction)

    ### B. SMART IMAGING LOGIC (HOLISTIC SEARCH)
    - **{{{{thorax_findings}}}}**: Look for **CECT Thorax**. If MISSING, look for **CHEST X-RAY (CXR)** or **HRCT Thorax**. Map those findings here instead of leaving blank.
    - **{{{{ncct_findings}}}}**: Look for **NCCT Brain**. If MISSING, look for **CT Head** or **CT Brain**.
    - **{{{{date_ncct}}}}**: Look specifically for the date written next to the scan title.
    - For NCCT Brain, MRI Brain, Bronchoscopy, USG Doppler, CECT:
    - Extract the text **EXACTLY** as written. Do not summarize.
    - If the report says "Normal study", write "Normal study".
    - Target Placeholders:
      * "{{date_ncct}}", "{{ncct_findings}}", "{{ncct_imp}}"
      * "{{date_mri}}", "{{mri_findings}}", "{{mri_imp}}"
      * "{{date_bronch}}", "{{bronch_findings}}", "{{bronch_imp}}"
      * "{{date_doppler}}", "{{doppler_findings}}", "{{doppler_imp}}"
      * "{{date_cect}}", "{{thorax_findings}}", "{{abdomen_findings}}", "{{cect_imp}}"

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

    try:
        # Run AI
        response = model.generate_content(prompt_content)

        # --- ADD THIS LINE ---
        # Strip 'models/' from the ID to match your PRICING keys (e.g. 'gemini-2.5-pro')
        clean_model_name = selected_model_id.replace("models/", "")
        log_usage(response, clean_model_name, note=f"Medical Extraction ({model_choice})")
        
        # Clean JSON (in case it added backticks)
        raw_text = response.text
        start_index = raw_text.find('{')
        end_index = raw_text.rfind('}') + 1
        
        if start_index != -1 and end_index != -1:
            json_text = raw_text[start_index:end_index]
            extracted_data = json.loads(json_text)

            # Fix: Check if grids are Strings and convert them to Dictionaries
            grid_keys = ["{{labs_json}}", "{{cardiac_json}}", "{{csf_json}}"]
            
            for key in grid_keys:
                # If the key exists AND it is a String (which causes the error)
                if key in extracted_data and isinstance(extracted_data[key], str):
                    try:
                        print(f"   -> Fixing stringified JSON for {key}...")
                        extracted_data[key] = json.loads(extracted_data[key])
                    except Exception as e:
                        print(f"   ⚠️ Warning: Could not fix {key}. It will be skipped. Error: {e}")
                        # Delete it so it doesn't crash the app later
                        del extracted_data[key]
        else:
            print("AI Raw Output:", raw_text) # For debugging
            return "Error: AI failed to generate JSON. Try again."
            
    except Exception as e:
        return f"AI Logic Error: {e}"

    # Determine filename
    patient_name = extracted_data.get("{{patient_name}}", "Unknown")
    patient_name = re.sub(r'[\\/*?:"<>|]', "", patient_name)
    
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
        return

    print(f"--- 4. Filling Data... ---")
    docs_service = build('docs', 'v1', credentials=creds)

    # --- 1. Fill Main Lab Grid (SAFE MODE) ---
    if "{{labs_json}}" in extracted_data:
        lab_data = extracted_data["{{labs_json}}"]
        # CHECK: Is it actually a dictionary?
        if isinstance(lab_data, dict):
            lab_requests = fill_smart_grid(docs_service, NEW_DOCUMENT_ID, lab_data, LAB_TEST_ORDER, "{{LAB_ANCHOR}}")
            if lab_requests:
                docs_service.documents().batchUpdate(documentId=NEW_DOCUMENT_ID, body={'requests': lab_requests}).execute(num_retries=5)
                print("   -> Main Lab Grid Filled.")
        else:
            print(f"   ⚠️ Skipping Labs: AI returned {type(lab_data)} instead of Dict")
        # Delete key so it doesn't break later steps
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
    
    # 1. Apply Defaults Loop
    for key, default_text in IMAGING_DEFAULTS.items():
        extracted_val = str(extracted_data.get(key, "")).strip()
        
        # LOGIC: If AI returns empty or "NOT_FOUND", use Professional Default.
        if not extracted_val or extracted_val.upper() in ["NOT_FOUND", "NONE", "", "NULL"]:
            final_data[key] = default_text
        else:
            final_data[key] = extracted_val

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
        "name": new_filename
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
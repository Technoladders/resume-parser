
import os
import json
import re
import logging
import datetime
import subprocess 
import io 
import ast
import traceback

# PDF/Office Processing
from pdf2image import convert_from_path, pdfinfo_from_path
from pdf2image.exceptions import PDFPageCountError
import pytesseract
from docx import Document
import textract 

# ReportLab / Platypus for PDF Generation
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image, Table, TableStyle, PageBreak, KeepTogether
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT
from reportlab.lib.units import inch
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter

# Async/Retry/Queueing
from tenacity import retry, stop_after_attempt, wait_exponential
from config import redis_conn # Assuming config.py defines redis_conn

# API Clients
from config import openai_client, supabase_admin_client as supabase, SUPABASE_STORAGE_BUCKET, SUPABASE_REPORT_PATH_PREFIX, SUPABASE_URL
# --- Custom Exception ---
class FileNotFoundInStorageError(Exception):
    """Custom exception for when a file is not found in Supabase Storage."""
    pass

# --- Logging Setup ---
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# --- ReportLab Styles ---
styles = getSampleStyleSheet()
# Custom Styles - Refined for Appeal
styles.add(ParagraphStyle(name='ReportTitle', parent=styles['h1'], alignment=TA_CENTER, fontSize=20, spaceBottom=24, textColor=colors.HexColor('#3A0CA3')))
styles.add(ParagraphStyle(name='SectionHeading', parent=styles['h2'], fontSize=14, spaceBefore=16, spaceAfter=8, textColor=colors.HexColor('#480CA8'), alignment=TA_LEFT))
styles.add(ParagraphStyle(name='SubHeading', parent=styles['h3'], fontSize=11, spaceBefore=10, spaceAfter=5, textColor=colors.HexColor('#480CA8'), fontName='Helvetica-Bold', alignment=TA_LEFT))
styles.add(ParagraphStyle(name='Body', parent=styles['Normal'], alignment=TA_LEFT, fontSize=10, leading=14, textColor=colors.darkslategray))
styles.add(ParagraphStyle(name='ListItem', parent=styles['Body'], leftIndent=20, spaceBefore=2, spaceAfter=2))
styles.add(ParagraphStyle(name='ScoreHighlight', parent=styles['Normal'], alignment=TA_RIGHT, fontSize=22, fontName='Helvetica-Bold', textColor=colors.HexColor('#3A0CA3')))
styles.add(ParagraphStyle(name='ScoreLabel', parent=styles['Normal'], alignment=TA_LEFT, fontSize=12, fontName='Helvetica-Bold', textColor=colors.HexColor('#4361EE')))
styles.add(ParagraphStyle(name='TableHeader', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=9, alignment=TA_LEFT, textColor=colors.white))
styles.add(ParagraphStyle(name='TableCell', parent=styles['Normal'], fontSize=9, leading=11))
styles.add(ParagraphStyle(name='WatermarkStyle', parent=styles['Normal'], alignment=TA_RIGHT, fontSize=8, textColor=colors.Color(0,0,0, alpha=0.15)))

# --- Helper Functions ---

def add_watermark(canvas, doc):
    """Adds 'hrumbles.ai' watermark to each page."""
    canvas.saveState()
    canvas.setFont('Helvetica', 8)
    canvas.setFillColor(colors.Color(0,0,0, alpha=0.15)) # Faint color
    # Position in top-right corner
    canvas.drawRightString(doc.pagesize[0] - 0.5*inch, doc.pagesize[1] - 0.5*inch, "hrumbles.ai")
    canvas.restoreState()
# --- END OF PASTE ---

def log_progress(job_id: str, step: str, message: str, data: dict = None):
    """Log a progress step and store in Redis."""
    try:
        log_entry = {
            "step": step,
            "message": message,
            "data": data or {},
            "timestamp": datetime.datetime.utcnow().isoformat() # Use UTC timestamp
        }
        logger.info(f"Job {job_id} - {step}: {message}")
        try:
            redis_conn.rpush(f"job_logs:{job_id}", json.dumps(log_entry))
        except Exception as redis_e:
            logger.error(f"Job {job_id} - REDIS_ERROR: Failed to push log to Redis for step '{step}': {redis_e}")
    except Exception as log_e:
        print(f"[FALLBACK LOG] Job {job_id} - {step}: {message} - Error in log_progress itself: {log_e}")

@retry(stop=stop_after_attempt(3), wait=wait_exponential(multiplier=1, min=4, max=10))
def download_resume(resume_path: str, job_id: str) -> str: # <--- MODIFICATION: Added job_id parameter
    """
    Downloads resume from Supabase Storage.
    Handles both full public URLs and relative paths.
    Raises FileNotFoundInStorageError if not found.
    """
    try:
        bucket_to_use = SUPABASE_STORAGE_BUCKET
        path_to_download = resume_path

        if resume_path.startswith("http"):
            try:
                full_path_in_bucket = resume_path.split('/public/')[1]
                path_parts = full_path_in_bucket.split('/')
                bucket_to_use = path_parts[0]
                path_to_download = '/'.join(path_parts[1:])
                # MODIFICATION: Now correctly uses job_id passed as a parameter
                log_progress(job_id, "download_resume_parse", "Parsed public URL", {"bucket": bucket_to_use, "path": path_to_download})
            except (IndexError, AttributeError):
                 log_progress(job_id, "download_resume_parse_error", "Could not parse public URL, using original path", {"url": resume_path})
                 pass
        
        response_content = supabase.storage.from_(bucket_to_use).download(path_to_download)
        
        local_path = f"/tmp/{os.path.basename(path_to_download)}"
        with open(local_path, "wb") as f:
            f.write(response_content)
        logger.info(f"Successfully downloaded resume to {local_path}")
        return local_path

    except Exception as e:
        error_message = str(e).lower()
        if 'not found' in error_message or 'status_code=400' in error_message or 'statuscode=400' in error_message or 'status_code=404' in error_message:
            logger.error(f"Resume file not found in Supabase Storage. Path attempted: '{path_to_download}' in bucket '{bucket_to_use}'. Error: {e}")
            raise FileNotFoundInStorageError(f"Resume file not found in Supabase Storage: {resume_path}") from e
        else:
            logger.error(f"Failed to download resume from Supabase. Path: {resume_path}. Error: {e}")
            raise Exception(f"Failed to download resume from Supabase: {str(e)}") from e

def get_pdf_page_count(pdf_path: str) -> int:
    """Gets page count using pdfinfo_from_path."""
    try:
        pdf_info = pdfinfo_from_path(pdf_path)
        count = int(pdf_info.get("Pages", 0)) # Use .get() for safety
        if count <= 0:
             raise ValueError("pdfinfo returned invalid page count")
        return count
    except (PDFPageCountError, ValueError, Exception) as e:
        logger.error(f"Failed to determine PDF page count for {pdf_path}: {e}")
        raise Exception(f"Failed to determine PDF page count: {str(e)}") from e

def convert_to_pdf(input_path: str, output_dir: str, job_id: str) -> str:
    """Converts DOC/DOCX to PDF using LibreOffice soffice command."""
    output_filename = os.path.splitext(os.path.basename(input_path))[0] + ".pdf"
    output_path = os.path.join(output_dir, output_filename)
    os.makedirs(output_dir, exist_ok=True)
    command = ['soffice', '--headless', '--convert-to', 'pdf', '--outdir', output_dir, input_path]
    try:
        log_progress(job_id, "convert_to_pdf_start", f"Converting {input_path} to PDF")
        result = subprocess.run(command, capture_output=True, text=True, timeout=120, check=True, encoding='utf-8', errors='replace')
        log_progress(job_id, "convert_to_pdf_success", f"Conversion command executed. Output: {output_path}", {"stdout": result.stdout[:500], "stderr": result.stderr[:500]})
        if not os.path.exists(output_path) or os.path.getsize(output_path) == 0:
             raise FileNotFoundError(f"LibreOffice conversion ran but output PDF not found or empty at {output_path}. Stderr: {result.stderr}")
        log_progress(job_id, "convert_to_pdf_complete", f"PDF conversion successful: {output_path}")
        return output_path
    except subprocess.CalledProcessError as e:
        log_progress(job_id, "convert_to_pdf_error", f"LibreOffice conversion failed (exit code {e.returncode})", {"stderr": e.stderr[:1000], "stdout": e.stdout[:1000]})
        raise Exception(f"LibreOffice conversion failed for {input_path}: {e.stderr}") from e
    except subprocess.TimeoutExpired as e:
        log_progress(job_id, "convert_to_pdf_error", f"LibreOffice conversion timed out for {input_path}")
        raise Exception(f"LibreOffice conversion timed out for {input_path}") from e
    except FileNotFoundError as e: # Catch specific file not found after conversion
        log_progress(job_id, "convert_to_pdf_error", f"Output PDF file missing after conversion command: {str(e)}")
        raise e # Re-raise the specific error
    except Exception as e:
        log_progress(job_id, "convert_to_pdf_error", f"Generic error during LibreOffice conversion for {input_path}: {str(e)}")
        raise Exception(f"Error during LibreOffice conversion: {str(e)}") from e

def extract_text_from_file(file_path: str, job_id: str) -> str:
    """Extracts text using Strategy 1: Convert DOC/DOCX to PDF, then OCR."""
    pdf_to_process = None
    temp_pdf_path = None
    text = ""
    try:
        file_extension = os.path.splitext(file_path)[1].lower()
        output_dir = "/tmp" # Directory for temporary PDFs

        if file_extension in [".doc", ".docx"]:
            log_progress(job_id, "extract_text_convert", f"Detected {file_extension}, attempting conversion to PDF.")
            try:
                temp_pdf_path = convert_to_pdf(file_path, output_dir, job_id)
                pdf_to_process = temp_pdf_path
                log_progress(job_id, "extract_text_convert_success", f"Converted to PDF: {pdf_to_process}")
            except Exception as conversion_err:
                log_progress(job_id, "extract_text_convert_failed", f"PDF conversion failed, attempting fallback extraction. Error: {conversion_err}")
                # Fallback Logic
                if file_extension == ".docx":
                     try:
                         doc = Document(file_path)
                         fallback_text_list = [p.text for p in doc.paragraphs if p.text.strip()]
                         text = "\n".join(fallback_text_list)
                         log_progress(job_id, "extract_text_fallback_docx", f"Used python-docx fallback, extracted {len(text)} chars.")
                         return text.strip() or "No text extracted (docx fallback)"
                     except Exception as docx_err:
                         log_progress(job_id, "extract_text_fallback_error", f"python-docx fallback failed: {docx_err}")
                         # Raise original conversion error if fallback also fails
                         raise Exception(f"Failed to convert DOCX to PDF and fallback also failed: {conversion_err}") from docx_err
                elif file_extension == ".doc":
                     try:
                         # Ensure textract and antiword (if needed) are installed
                         text = textract.process(file_path).decode('utf-8', errors='ignore')
                         log_progress(job_id, "extract_text_fallback_doc", f"Used textract fallback, extracted {len(text)} chars.")
                         return text.strip() or "No text extracted (doc fallback)"
                     except Exception as textract_err:
                          log_progress(job_id, "extract_text_fallback_error", f"textract fallback failed for .doc: {textract_err}")
                          raise Exception(f"Failed to convert DOC to PDF and fallback extraction also failed: {conversion_err}") from textract_err
        elif file_extension == ".pdf":
            pdf_to_process = file_path # Use the original PDF
        else:
            raise Exception(f"Unsupported file format: {file_extension}")

        # --- Process the PDF (Original or Converted) using OCR ---
        if pdf_to_process:
            log_progress(job_id, "extract_text_ocr_start", f"Processing PDF with OCR: {os.path.basename(pdf_to_process)}")
            try:
                total_pages = get_pdf_page_count(pdf_to_process)
                log_progress(job_id, "extract_text_page_count", f"PDF has {total_pages} pages")
                custom_config = r'--oem 3 --psm 6' # Adjust psm if needed

                all_page_texts = []
                for page_num in range(1, total_pages + 1):
                    page_log_prefix = f"extract_text_page_{page_num}"
                    log_progress(job_id, page_log_prefix, f"Processing page {page_num}/{total_pages}")
                    try:
                        images = convert_from_path(pdf_to_process, dpi=200, first_page=page_num, last_page=page_num, timeout=60)
                        if not images:
                            log_progress(job_id, f"{page_log_prefix}_warn", "No image generated")
                            continue
                        page_text = pytesseract.image_to_string(images[0], config=custom_config)
                        all_page_texts.append(page_text)
                        log_progress(job_id, f"{page_log_prefix}_success", f"Extracted text", {"length": len(page_text)})
                        del images # Clean up memory
                    except Exception as ocr_page_err:
                         log_progress(job_id, f"{page_log_prefix}_error", f"Error during OCR: {ocr_page_err}")
                         all_page_texts.append(f"\n--- Error processing page {page_num} ---") # Add error marker

                text = "\n--- Page Break ---\n".join(all_page_texts) # Join pages with separator
                log_progress(job_id, "extract_text_ocr_complete", "Finished OCR processing")

            except Exception as pdf_processing_err:
                 log_progress(job_id, "extract_text_ocr_error", f"Error processing PDF {pdf_to_process}: {pdf_processing_err}")
                 raise Exception(f"Error processing PDF {pdf_to_process}: {str(pdf_processing_err)}") from pdf_processing_err
        else:
             # Should only happen if conversion failed AND fallback failed above
             raise Exception("No processable PDF found after conversion attempts.")

        log_progress(job_id, "extract_text_final", "Text extraction completed.", {"final_text_length": len(text)})
        return text.strip() or "No text extracted"

    except Exception as e:
        tb_str = traceback.format_exc()
        log_progress(job_id, "extract_text_fatal_error", f"Fatal error during text extraction: {str(e)}", {"traceback": tb_str})
        raise Exception(f"Failed to extract text from file {file_path}: {str(e)}") from e
    finally:
         # Clean up temporary PDF if created
         if temp_pdf_path and os.path.exists(temp_pdf_path):
             try:
                 os.remove(temp_pdf_path)
                 log_progress(job_id, "extract_text_cleanup", f"Removed temporary PDF: {temp_pdf_path}")
             except Exception as cleanup_err:
                 log_progress(job_id, "extract_text_cleanup_error", f"Failed to remove temp PDF {temp_pdf_path}: {cleanup_err}")

def clean_ai_output(text: str) -> str:
    """Cleans AI output, extracts JSON."""
    text = text.strip()
    if text.startswith('```json'): text = text[7:].rstrip('```')
    elif text.startswith('```'): text = text[3:].rstrip('```')
    text = text.strip()
    start = text.find('{'); end = text.rfind('}')
    if start != -1 and end != -1 and end > start: text = text[start:end+1]
    else: logger.warning(f"clean_ai_output: No JSON object markers found in preview: {text[:200]}")
    text = re.sub(r'[\x00-\x1F\x7F]', '', text) # Remove control chars
    text = re.sub(r',\s*([}\]])', r'\1', text) # Remove trailing commas
    return text

@retry(stop=stop_after_attempt(3), wait=wait_exponential(multiplier=1, min=4, max=10))
def generate_report(resume_text: str, job_description: str, job_id: str, organization_id: str, user_id: str, analysis_config: dict = None) -> dict:
    """Generates analysis report using OpenAI, includes detailed logging."""
    attempt_id = os.urandom(4).hex()
    current_step = "init"

     # --- START: MODIFICATION ---
    input_tokens = 0
    output_tokens = 0
    status = 'FAILURE'
    analysis_for_log = None
    raw_ai_text = ""
    # --- END: MODIFICATION ---
    
    # --- DYNAMIC PROMPT GENERATION ---
    
    # Default Defaults
    weights = {
        "Technical Skills": 45,
        "Work Experience": 30,
        "Projects": 15,
        "Education": 10
    }
    
    include_soft = False
    include_achievements = False

    # Override if config exists
    if analysis_config:
        if "weights" in analysis_config:
            weights = analysis_config["weights"]
        if "sections" in analysis_config:
            include_soft = analysis_config["sections"].get("Soft Skills", False)
            include_achievements = analysis_config["sections"].get("Achievements", False)

    # Build the weightage string dynamically
    weightage_text = "4. WEIGHTAGE RULES (STRICT):\n   Use these EXACT weights provided by the user configuration:\n"
    for section, value in weights.items():
        weightage_text += f"     - {section} = {value}%\n"
    
    weightage_text += "\n   Note: If a section has 0% weight, it must NOT affect the overall score, but you may still analyze it for qualitative feedback.\n"

    # Build Section JSON structure dynamically based on preferences
    sections_json_structure = ""
    
    # Required Sections
    sections_json_structure += """
    {{
      "section": "Technical Skills",
      "weightage": {tech_w},
      "submenus": [
        {{ "submenu": "Core Required Skills", "weightage": 100, "score": number, "weighted_score": number, "remarks": "..." }}
      ],
      "section_contribution": number
    }},
    {{
      "section": "Work Experience",
      "weightage": {exp_w},
      "submenus": [
        {{ "submenu": "Role Relevance", "weightage": 50, "score": number, "weighted_score": number, "remarks": "..." }},
        {{ "submenu": "Experience Depth", "weightage": 30, "score": number, "weighted_score": number, "remarks": "..." }},
        {{ "submenu": "Recency & Continuity", "weightage": 20, "score": number, "weighted_score": number, "remarks": "..." }}
      ],
      "section_contribution": number
    }},
    {{
      "section": "Projects",
      "weightage": {proj_w},
      "submenus": [
         {{ "submenu": "Relevance & Complexity", "weightage": 100, "score": number, "weighted_score": number, "remarks": "..." }}
      ],
      "section_contribution": number
    }},
    {{
      "section": "Education",
      "weightage": {edu_w},
      "submenus": [
        {{ "submenu": "Degree Alignment", "weightage": 100, "score": number, "weighted_score": number, "remarks": "..." }}
      ],
      "section_contribution": number
    }},
    """.format(
        tech_w=weights.get("Technical Skills", 0),
        exp_w=weights.get("Work Experience", 0),
        proj_w=weights.get("Projects", 0),
        edu_w=weights.get("Education", 0)
    )

    # Optional Sections
    if include_soft:
        sections_json_structure += """
        {{
          "section": "Soft Skills",
          "weightage": {soft_w},
          "submenus": [{{ "submenu": "Communication & Leadership", "weightage": 100, "score": number, "weighted_score": number, "remarks": "..." }}],
          "section_contribution": number
        }},
        """.format(soft_w=weights.get("Soft Skills", 0))
    else:
        sections_json_structure += """
        {{
          "section": "Soft Skills",
          "weightage": 0,
          "submenus": [],
          "remarks": "Inferred soft skills only (Not Scored)"
        }},
        """

    if include_achievements:
        sections_json_structure += """
        {{
          "section": "Achievements",
          "weightage": {ach_w},
          "submenus": [{{ "submenu": "Awards & Certifications", "weightage": 100, "score": number, "weighted_score": number, "remarks": "..." }}],
          "section_contribution": number
        }}
        """.format(ach_w=weights.get("Achievements", 0))
    else:
         sections_json_structure += """
        {{
          "section": "Achievements",
          "weightage": 0,
          "submenus": [],
          "remarks": "Certifications and awards only (Not Scored)"
        }}
        """
    
    try:
        current_step = "log_start"; log_progress(job_id, f"generate_report_start_{attempt_id}", "Entering generate_report")
        current_step = "log_raw_inputs"; log_progress(job_id, f"generate_report_raw_inputs_{attempt_id}", "Raw input previews", {"resume_preview": resume_text[:200], "jd_preview": job_description[:200]})
        current_step = "log_input_preview"; log_progress(job_id, f"generate_report_input_{attempt_id}", "Preparing prompt", {"resume_preview": resume_text[:1000], "jd_preview": job_description[:1000]})
        
         # --- FIXED PROMPT TEMPLATE (Double Curly Braces) ---
        prompt_template = """
You are an expert ATS system analyzing candidate fit.
Return ONLY a valid JSON object. No explanations. No markdown.

=== ANALYSIS FRAMEWORK ===

1. MATCHING METHODOLOGY (STRICT):
   - Explicit Match (✅ score 8–10): Direct, current evidence in resume
   - Strong Implied (⚠ score 6–7): Clear indirect or strongly related evidence
   - Weak Implied (⚠ score 4–5): Transferable, limited, or dated evidence
   - No Match (❌ score 0–3): Absent or insufficient evidence

2. SCORING SCALE (CRITICAL – FOLLOW EXACTLY):
   - ALL scores are on a 0–10 scale until the FINAL step
   - NEVER divide scores by 10
   - NEVER normalize scores early
   - Multiply by 10 ONLY ONCE at the very end to get a 0–100 score

3. SECTION SCORE CALCULATION (LOCKED):
   For each section:
     raw_score = Σ(submenu.score × submenu.weightage / 100)
     section_contribution = raw_score × (section.weightage / 100)

   Final calculation (ONLY here):
     overall_match_score = Σ(section_contribution) × 10

4. {weightage_rules_block}

5. CONFIDENCE & HIRING LOGIC:
   - If ALL must-have skills are met → confidence_level = "high"
   - Nice-to-have gaps must NOT reduce confidence
   - Critical gaps ONLY if JD explicitly marks skill as MUST

=== INPUT DATA ===
Job Description:
{job_description}

Resume:
{resume_text}

=== REQUIRED JSON STRUCTURE ===

{{
  "candidate_info": {{
    "name": "string or 'Unknown'",
    "email": "string or ''",
    "phone": "string or ''",
    "linkedin": "string or ''",
    "github": "string or ''"
  }},

  "overall_match_score": number,

  "match_quality": {{
    "hiring_recommendation": "strong_yes | yes | maybe | no",
    "confidence_level": "high | medium | low",
    "summary": "2–3 concise, evidence-based sentences on candidate fit",
    "key_differentiators": ["array of unique strengths"]
  }},

  "requirements_coverage": {{
    "must_have_skills_met": "X/Y",
    "nice_to_have_skills_met": "X/Y",
    "critical_gaps": ["array or empty"]
  }},

  "matched_skills": [
    {{
      "requirement": "Exact skill from JD",
      "matched": "yes | partial | no",
      "score": number,
      "evidence": "Exact resume evidence or 'Not mentioned'",
      "recency": "current | dated | unknown"
    }}
  ],

  "experience_analysis": {{
    "total_years": number,
    "relevant_years": number,
    "role_progression": "ascending | lateral | descending",
    "companies": [
      {{
        "name": "string",
        "designation": "string or '-'",
        "duration": "YYYY-MM - YYYY-MM or '-'",
        "relevance_score": number
      }}
    ]
  }},

  "section_wise_scoring": [
    {sections_json_block}
  ],

  "top_skills": ["top 5–7 strongest skills"],
  "development_gaps": ["skills to improve"],
  "additional_certifications": ["certifications not required by JD"],
  "red_flags": ["array or empty"],

  "resume_quality": {{
    "parsing_confidence": number,
    "format_issues": ["array or empty"],
    "completeness_score": number
  }}
}}

=== FINAL VALIDATION (MANDATORY) ===
Before output:
- Section weightages MUST sum to 100
- overall_match_score MUST equal Σ(section_contribution) × 10
- No section with 0% weightage may influence score
- Do NOT divide any score by 10
- Output ONLY valid JSON

""" # Keep prompt concise here, assume details are known by model or adjust if needed
        current_step = "before_format"; log_progress(job_id, f"generate_report_before_format_{attempt_id}", "Formatting prompt")
        prompt = prompt_template.format(job_description=job_description, resume_text=resume_text, weightage_rules_block=weightage_text, sections_json_block=sections_json_structure)
        current_step = "after_format"; log_progress(job_id, f"generate_report_after_format_{attempt_id}", "Prompt formatted", {"prompt_preview": prompt[:500]})

        current_step = "before_api_call"; log_progress(job_id, f"generate_report_before_api_call_{attempt_id}", "Calling OpenAI API")
        
        # 2. NEW API CALL PARAMETERS
        # Note: "gpt-4.1" does not exist in public API. Assuming "gpt-4o" or "gpt-4-turbo".
        # Using "gpt-4o" as it is the current flagship.
        response = openai_client.chat.completions.create(
            model="gpt-4o", 
            messages=[{"role": "user", "content": prompt}],
            response_format={"type": "json_object"},
            temperature=1,
            top_p=0.80,
        )
        
        current_step = "after_api_call"; log_progress(job_id, f"generate_report_after_api_call_{attempt_id}", "Returned from OpenAI API")

        if response.usage:
            input_tokens = response.usage.prompt_tokens
            output_tokens = response.usage.completion_tokens
        
        raw_ai_text = response.choices[0].message.content if response.choices else ""

        current_step = "log_raw_output"; log_progress(job_id, f"generate_report_raw_output_{attempt_id}", "Raw OpenAI response", {"preview": raw_ai_text[:2000]})
        current_step = "clean_output"; ai_output = clean_ai_output(raw_ai_text); log_progress(job_id, f"generate_report_cleaned_output_{attempt_id}", "Cleaned OpenAI response", {"preview": ai_output[:2000]})

        current_step = "parse_json"
        try:
            raw_report = json.loads(ai_output)
        except json.JSONDecodeError as e:
            log_progress(job_id, f"generate_report_parse_error_{attempt_id}", f"Failed JSON parse: {e}", {"preview": ai_output[:2000]})
            raise Exception(f"Invalid JSON from OpenAI. Original error: {e}") from e

        status = 'SUCCESS'
        
        # 3. BACKWARD COMPATIBILITY MAPPING (Crucial Step)
        # This converts the new nested structure to the flat structure your PDF/DB code expects.
        report = raw_report.copy()
        
         # 1. Map Candidate Info to top level for PDF
        c_info = raw_report.get("candidate_info", {})
        report["candidate_name"] = c_info.get("name", "Unknown")
        report["email"] = c_info.get("email", "")
        report["phone_number"] = c_info.get("phone", "")
        report["linkedin"] = c_info.get("linkedin", "")
        report["github"] = c_info.get("github", "")

        # 2. Map Summary (The new prompt puts summary inside 'match_quality')
        if "match_quality" in raw_report and "summary" in raw_report["match_quality"]:
            report["summary"] = raw_report["match_quality"]["summary"]

        # Map Companies
        # New format: experience_analysis -> companies -> duration
        # Old format: companies -> years
        exp_analysis = raw_report.get("experience_analysis", {})
        companies_raw = exp_analysis.get("companies", [])
        mapped_companies = []
        for c in companies_raw:
            mapped_companies.append({
                "name": c.get("name"),
                "designation": c.get("designation"),
                "years": c.get("duration", "-") 
            })
        report["companies"] = mapped_companies
        
        # --- 4. Gaps & Coverage (CRITICAL FIX) ---
        # Prioritize 'critical_gaps' from requirements_coverage if available
        req_coverage = raw_report.get("requirements_coverage", {})
        critical_gaps = req_coverage.get("critical_gaps", [])
        dev_gaps = raw_report.get("development_gaps", [])
        
        # Consolidate into the field the DB expects
        if critical_gaps:
            report["missing_or_weak_areas"] = critical_gaps
        elif dev_gaps:
             report["missing_or_weak_areas"] = dev_gaps
        else:
             report["missing_or_weak_areas"] = []

        # Ensure these are always present for DB
        report["red_flags"] = raw_report.get("red_flags", [])
        report["requirements_coverage"] = req_coverage
        
        if "matched_skills" in report:
            for skill in report["matched_skills"]:
                # Map 'evidence' to 'details' for PDF compatibility
                if "evidence" in skill and "details" not in skill:
                    skill["details"] = skill["evidence"]

        analysis_for_log = report

        current_step = "log_parsed_report"; log_progress(job_id, f"generate_report_parsed_{attempt_id}", "Parsed report structure", {"keys": list(report.keys())})

        current_step = "validate_report"
        # Updated validation list to check for the FLAT keys we just mapped
        required_fields = ["overall_match_score", "matched_skills", "summary", "companies", "missing_or_weak_areas", "top_skills", "development_gaps", "section_wise_scoring", "candidate_name", "email"]
        for field in required_fields:
            if field not in report: 
                # If missing, try to patch it to avoid crash
                logger.warning(f"Missing field {field} in report, adding default.")
                report[field] = [] if field in ["companies", "missing_or_weak_areas", "top_skills", "development_gaps"] else "N/A"


        current_step = "normalize_companies"
        # Company normalization logic (ensure normalize_company_name is defined)
        if "companies" in report and isinstance(report["companies"], list):
            processed_companies = []
            for company in report["companies"]:
                if isinstance(company, dict) and "name" in company:
                    processed_companies.append({"name": company["name"], "normalized_name_for_dedup": normalize_company_name(company["name"]), "designation": company.get("designation", "-"), "years": company.get("years", "-")})
            unique_companies_final = []
            seen_normalized_names = set()
            for company in reversed(processed_companies):
                norm_name = company["normalized_name_for_dedup"]
                if norm_name not in seen_normalized_names:
                    del company["normalized_name_for_dedup"]; unique_companies_final.append(company); seen_normalized_names.add(norm_name)
            report["companies"] = list(reversed(unique_companies_final))
        else: report["companies"] = []; log_progress(job_id, f"generate_report_norm_warn_{attempt_id}", "Companies missing/not list")

        current_step = "log_final_report"
        try: report_preview = json.dumps(report)[:2000]
        except Exception: report_preview = "Error creating preview"
        log_progress(job_id, f"generate_report_final_parsed_{attempt_id}", "Final processed report preview", {"preview": report_preview})

        current_step = "final_success"; log_progress(job_id, f"generate_report_success_{attempt_id}", "Report generated successfully")
        return report

    except Exception as e:
        safe_raw_text = json.dumps(raw_ai_text) # Safely encode the raw text as a JSON string
        analysis_for_log = {
            "error": str(e), 
            "failed_step": current_step, 
            "raw_response": safe_raw_text
        }
        tb_str = traceback.format_exc()
        logger.error(f"Job {job_id} - generate_report_error_{attempt_id}: Exception at step '{current_step}': {e}\n{tb_str}")
        log_progress(job_id, f"generate_report_error_{attempt_id}", f"Exception at step '{current_step}': {e}", {"type": type(e).__name__, "step": current_step, "traceback": tb_str})
        raise Exception(f"Failed report generation at step '{current_step}': {e}") from e
    
    finally:
        try:
            if organization_id and user_id:
                log_progress(job_id, f"log_ai_usage_{attempt_id}", "Logging AI usage to database")
                supabase.from_('hr_gemini_usage_log').insert({
                    "organization_id": organization_id,
                    "created_by": user_id,
                    "status": status,
                    "input_tokens": input_tokens,
                    "output_tokens": output_tokens,
                    "analysis_response": analysis_for_log,
                    "parsed_email": analysis_for_log.get('email') if isinstance(analysis_for_log, dict) else None,
                     "usage_type": 'resume_validation_openai'
                }).execute()
            else:
                log_progress(job_id, f"log_ai_usage_skipped_{attempt_id}", "Skipping usage logging due to missing org/user ID")
        except Exception as log_db_e:
            logger.error(f"Job {job_id} - FAILED_TO_LOG_USAGE: Could not write to hr_gemini_usage_log. Error: {log_db_e}")
            log_progress(job_id, f"log_ai_usage_error_{attempt_id}", "Failed to log usage to DB", {"error": str(log_db_e)})
        

def save_report_as_pdf(report: dict, output_path: str, job_id: str):
    """Saves the analysis report as a styled PDF using Platypus."""
    try:
        log_progress(job_id, "pdf_generation_start", "Starting PDF generation with Platypus")
        doc = SimpleDocTemplate(output_path, pagesize=letter, leftMargin=0.75*inch, rightMargin=0.75*inch, topMargin=1.0*inch, bottomMargin=0.75*inch)
        story = []; bullet = '•'

        # --- Build Story ---
        story.append(Paragraph("Resume Analysis Report", styles['ReportTitle']))

        # Candidate Details Section
        candidate_section = [Paragraph("Candidate Details", styles['SectionHeading'])]
        candidate_data = [ [Paragraph(f"<b>{k.replace('_',' ').title()}:</b>", styles['TableCell']), Paragraph(str(report.get(k, 'N/A')), styles['TableCell'])] for k in ['candidate_name', 'email', 'phone_number', 'linkedin', 'github'] ]
        candidate_table = Table(candidate_data, colWidths=[1.2*inch, 5.8*inch], style=[ ('VALIGN', (0,0), (-1,-1), 'TOP'), ('LEFTPADDING', (0,0), (-1,-1), 0), ('BOTTOMPADDING', (0,0), (-1,-1), 4) ])
        candidate_section.append(candidate_table); candidate_section.append(Spacer(1, 0.25*inch)); story.append(KeepTogether(candidate_section))

        # Score Section
        score_section = []
        score_data = [[ Paragraph("Overall Match Score", styles['ScoreLabel']), Paragraph(f"{report.get('overall_match_score', 0)}%", styles['ScoreHighlight']) ]]
        score_table = Table(score_data, colWidths=[5*inch, 2*inch], style=[ ('VALIGN', (0,0), (-1,-1), 'MIDDLE'), ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#F0F3FF')), ('LEFTPADDING', (0,0), (0,0), 12), ('RIGHTPADDING', (1,0), (1,0), 12), ('TOPPADDING', (0,0), (-1,-1), 10), ('BOTTOMPADDING', (0,0), (-1,-1), 10), ('LINEBELOW', (0,0), (-1,0), 1, colors.HexColor('#4361EE')) ])
        score_section.append(score_table); score_section.append(Spacer(1, 0.25*inch)); story.append(KeepTogether(score_section))

      
        summary_section = [ Paragraph("Overall Summary", styles['SectionHeading']), Paragraph(report.get('summary', "N/A"), styles['Body']), Spacer(1, 0.25*inch) ]
        story.append(KeepTogether(summary_section))

     
        skills_section = [Paragraph("Skills Overview", styles['SectionHeading'])]
        top_skills_list = report.get("top_skills", []); missed_skills_list = report.get("missing_or_weak_areas", [])
        top_flowables = [Paragraph("<b>Top Skills</b>", styles['SubHeading'])] + ([Paragraph(f"{bullet} {s}", styles['ListItem']) for s in top_skills_list] if top_skills_list else [Paragraph("N/A", styles['Body'])])
        missed_flowables = [Paragraph("<b>Missed / Weak Areas</b>", styles['SubHeading'])] + ([Paragraph(f"{bullet} {a}", styles['ListItem']) for a in missed_skills_list] if missed_skills_list else [Paragraph("N/A", styles['Body'])])
        skills_data = [[top_flowables, missed_flowables]] # Removed KeepTogether from here as it caused issues
        skills_table = Table(skills_data, colWidths=[3.5*inch, 3.5*inch], style=[ ('VALIGN', (0,0), (-1,-1), 'TOP'), ('BOX', (0,0), (-1,-1), 1, colors.HexColor('#D9E2FF')), ('LEFTPADDING', (0,0), (-1,-1), 12), ('RIGHTPADDING', (0,0), (-1,-1), 12), ('TOPPADDING', (0,0), (-1,-1), 8), ('BOTTOMPADDING', (0,0), (-1,-1), 8) ])
        skills_section.append(skills_table); skills_section.append(Spacer(1, 0.25*inch)); story.append(KeepTogether(skills_section)) # Keep heading and table together

     
        matched_skills_section = [Paragraph("Detailed Skill Match", styles['SectionHeading'])]
        matched_skills_list = report.get("matched_skills", [])
        if matched_skills_list:
            hdr = [Paragraph(h, styles['TableHeader']) for h in ["Requirement", "Match", "Evidence / Details"]]
            matched_data = [hdr] + [ [Paragraph(s.get('requirement','N/A'), styles['TableCell']), Paragraph({"yes": "✅ Yes", "partial": "⚠️ Partial", "no": "❌ No"}.get(s.get('matched', 'no'), "❓"), styles['TableCell']), Paragraph(s.get('details','N/A'), styles['TableCell'])] for s in matched_skills_list ]
            matched_table = Table(matched_data, colWidths=[2.5*inch, 0.8*inch, 3.7*inch], style=[ ('BACKGROUND',(0,0),(-1,0),colors.HexColor('#4361EE')), ('TEXTCOLOR',(0,0),(-1,0),colors.white), ('ALIGN',(0,0),(-1,-1),'LEFT'), ('VALIGN',(0,0),(-1,-1),'TOP'), ('FONTNAME',(0,0),(-1,0),'Helvetica-Bold'), ('BOTTOMPADDING',(0,0),(-1,0),8), ('TOPPADDING',(0,0),(-1,0),6), ('GRID',(0,0),(-1,-1),0.5,colors.lightgrey), ('LEFTPADDING',(0,0),(-1,-1),6), ('RIGHTPADDING',(0,0),(-1,-1),6), ('TOPPADDING',(0,1),(-1,-1),4), ('BOTTOMPADDING',(0,1),(-1,-1),4) ])
            matched_skills_section.append(matched_table)
        else: matched_skills_section.append(Paragraph("N/A", styles['Body']))
        matched_skills_section.append(Spacer(1, 0.25*inch)); story.append(KeepTogether(matched_skills_section))

        company_section = [Paragraph("Companies Mentioned", styles['SectionHeading'])]
        companies = report.get("companies", [])
        if companies: company_section.extend([Paragraph(f"{bullet} {c.get('name','N/A')} ({c.get('designation','-')}, {c.get('years','-')})", styles['ListItem']) for c in companies])
        else: company_section.append(Paragraph("N/A", styles['Body']))
        company_section.append(Spacer(1, 0.25*inch)); story.append(KeepTogether(company_section))

       
        gaps_section = [Paragraph("Development Gaps", styles['SectionHeading'])]
        gaps = report.get("development_gaps", [])
        if gaps: gaps_section.extend([Paragraph(f"{bullet} {g}", styles['ListItem']) for g in gaps])
        else: gaps_section.append(Paragraph("N/A", styles['Body']))
        gaps_section.append(Spacer(1, 0.25*inch)); story.append(KeepTogether(gaps_section))

     
        certs_section = [Paragraph("Additional Certifications", styles['SectionHeading'])]
        certs = report.get("additional_certifications", [])
        if certs: certs_section.extend([Paragraph(f"{bullet} {c}", styles['ListItem']) for c in certs])
        else: certs_section.append(Paragraph("N/A", styles['Body']))
        certs_section.append(Spacer(1, 0.25*inch)); story.append(KeepTogether(certs_section))

       
        scoring_section_content = [Paragraph("Section-Wise Scoring", styles['SectionHeading'])]
        scoring_list = report.get("section_wise_scoring", [])
        if scoring_list:
            hdr = [Paragraph(h, styles['TableHeader']) for h in ["Section", "Sub-Section", "Score", "Remarks"]]
            scoring_data = [hdr]
            for section in scoring_list:
                first_row = True
                for submenu in section.get("submenus", []):
                    sec_disp = Paragraph(f"<b>{section.get('section','N/A')}</b> ({section.get('weightage',0)}%)", styles['TableCell']) if first_row else ""
                    scoring_data.append([ sec_disp, Paragraph(submenu.get('submenu','N/A'), styles['TableCell']), Paragraph(f"{submenu.get('score',0)}/10", styles['TableCell']), Paragraph(submenu.get('remarks',''), styles['TableCell']) ])
                    first_row = False
            scoring_table = Table(scoring_data, colWidths=[1.5*inch, 1.7*inch, 0.8*inch, 3.0*inch], style=[ ('BACKGROUND',(0,0),(-1,0),colors.HexColor('#4361EE')), ('TEXTCOLOR',(0,0),(-1,0),colors.white), ('ALIGN',(0,0),(-1,-1),'LEFT'), ('VALIGN',(0,0),(-1,-1),'TOP'), ('FONTNAME',(0,0),(-1,0),'Helvetica-Bold'), ('BOTTOMPADDING',(0,0),(-1,0),8), ('TOPPADDING',(0,0),(-1,0),6), ('GRID',(0,0),(-1,-1),0.5,colors.lightgrey), ('LEFTPADDING',(0,0),(-1,-1),6), ('RIGHTPADDING',(0,0),(-1,-1),6), ('TOPPADDING',(0,1),(-1,-1),4), ('BOTTOMPADDING',(0,1),(-1,-1),4) ])
            scoring_section_content.append(scoring_table)
        else: scoring_section_content.append(Paragraph("N/A", styles['Body']))
        scoring_section_content.append(Spacer(1, 0.25*inch)); story.append(KeepTogether(scoring_section_content))

        # --- Build PDF ---
        log_progress(job_id, "pdf_generation_build", "Building PDF document")
        doc.build(story, onFirstPage=add_watermark, onLaterPages=add_watermark)
        log_progress(job_id, "pdf_generation_success", "PDF generated successfully")

    except Exception as e:
        tb_str = traceback.format_exc()
        logger.error(f"Job {job_id} - Failed PDF generation: {e}\n{tb_str}")
        log_progress(job_id, "pdf_generation_error", f"Failed PDF generation: {e}", {"type": type(e).__name__, "traceback": tb_str})
        raise Exception(f"Failed to save report as PDF: {e}") from e

@retry(stop=stop_after_attempt(3), wait=wait_exponential(multiplier=1, min=4, max=10))
def upload_report(report_path: str, destination_path: str, job_id: str):
    """Uploads the generated report PDF to Supabase Storage."""
    try:
        file_exists = os.path.exists(report_path); file_size = os.path.getsize(report_path) if file_exists else -1
        log_progress(job_id, "upload_report_info", "Preparing report upload", {"local": report_path, "dest": destination_path, "exists": file_exists, "size": file_size})
        if not file_exists or file_size <= 0: raise Exception(f"Local report file missing or empty: {report_path}")
        with open(report_path, "rb") as f:
            response = supabase.storage.from_(SUPABASE_STORAGE_BUCKET).upload(
                path=destination_path, file=f, file_options={"upsert": "true", "content-type": "application/pdf"}
            )
        log_progress(job_id, "upload_report_success", "Upload API call finished", {"destination": destination_path})
      
        return {"message": "Upload successful"} 
    except Exception as e:
        tb_str = traceback.format_exc(); error_type = type(e).__name__
        log_progress(job_id, "upload_report_exception", f"Upload failed ({error_type}): {e}", {"local": report_path, "dest": destination_path, "type": error_type, "traceback": tb_str})
        raise Exception(f"Failed upload ({error_type}): {e}") from e

def normalize_company_name(name: str) -> str:
    """Normalizes company name."""
    if not name: return ""
    normalized = name.lower().strip()
    normalized = re.sub(r'\s*(ltd|limited|inc|corp|corporation|llc|co)\.?\s*$', '', normalized, flags=re.IGNORECASE)
    normalized = re.sub(r'[^\w\s-]', '', normalized) 
    return ' '.join(normalized.split())


def process_analysis(job_id: str, candidate_id: str, resume_path: str, job_description_from_request: str, organization_id: str, user_id: str, analysis_config: dict = None):
    """Main background task to process resume analysis."""
    logger.info(f"Starting process_analysis for job_id: {job_id}, candidate_id: {candidate_id}, organization_id: {organization_id}, user_id: {user_id}")
    local_resume_path = None
    local_report_path = None

    try:
        
        log_progress(job_id, "init", "Task started", {"candidate_id": candidate_id, "resume_path": resume_path, "organization_id": organization_id})

    
        log_progress(job_id, "download_resume", f"Attempting download: {resume_path}")
        try:
            local_resume_path = download_resume(resume_path, job_id)
            log_progress(job_id, "download_resume_success", "Resume downloaded", {"local": local_resume_path, "size": os.path.getsize(local_resume_path) if local_resume_path and os.path.exists(local_resume_path) else -1})
        except FileNotFoundInStorageError as resume_not_found_err:
            log_progress(job_id, "error_resume_not_found", f"STOPPING TASK - Resume not found: {resume_path}. No DB changes.")
            logger.error(f"Job {job_id} - STOPPING TASK - Resume not found: {resume_path} - {resume_not_found_err}")
            return {"status": "failed", "candidate_id": candidate_id, "error": str(resume_not_found_err)}

        # Step 2: Validate job_id exists and fetch description
        log_progress(job_id, "fetch_jd", f"Fetching description for job_id {job_id} and organization {organization_id}")
        job_response = supabase.table("hr_jobs").select("description").eq("id", job_id).eq("organization_id", organization_id).execute()
        if not job_response.data or not job_response.data[0].get("description"): raise Exception(f"Job description not found for job {job_id} in organization {organization_id}")
        job_description_from_db = job_response.data[0]["description"]; log_progress(job_id, "fetch_jd_success", "Fetched job description")

        # Step 3: Extract text
        log_progress(job_id, "extract_text", "Extracting text from resume")
        resume_text = extract_text_from_file(local_resume_path, job_id)
        log_progress(job_id, "extract_text_success", "Text extracted", {"length": len(resume_text)})

      
        log_progress(job_id, "generate_report", "Generating report with OpenAI")
        report = generate_report(resume_text, job_description_from_db, job_id, organization_id, user_id, analysis_config)
        log_progress(job_id, "generate_report_success", "Report generated", {"score": report.get("overall_match_score", "N/A")})

      
        log_progress(job_id, "check_candidate", f"Checking/creating candidate {candidate_id}")
        candidate_check_resp = supabase.table("hr_candidates").select("id").eq("id", candidate_id).eq("organization_id", organization_id).execute()
        if not candidate_check_resp.data:
             log_progress(job_id, "create_candidate", f"Candidate {candidate_id} not found, creating...")
             insert_data = { "id": candidate_id, "name": report.get("candidate_name", "Unknown"), "email": report.get("email") or f"unknown_{candidate_id}@example.com", "phone_number": report.get("phone_number"), "linkedin_url": report.get("linkedin"), "github_url": report.get("github"), "organization_id": organization_id }
             supabase.table("hr_candidates").insert(insert_data).execute()
             log_progress(job_id, "create_candidate_success", f"Created new candidate record")

        
        log_progress(job_id, "save_report", "Saving report as PDF")
        report_filename = f"report_{job_id}_{candidate_id}.pdf"; local_report_path = f"/tmp/{report_filename}"
        save_report_as_pdf(report, local_report_path, job_id)
        log_progress(job_id, "save_report_success", "Report saved successfully")

        
        log_progress(job_id, "upload_report", "Uploading report to Supabase Storage")
        report_destination_path = f"{SUPABASE_REPORT_PATH_PREFIX}/{job_id}/{report_filename}"
        upload_report(local_report_path, report_destination_path, job_id)
        log_progress(job_id, "upload_report_success", "Report uploaded successfully")

        log_progress(job_id, "process_companies", "Processing company associations")
        company_entries = []
        raw_companies = report.get("companies", [])
        if isinstance(raw_companies, list):
            for company in raw_companies:
                if isinstance(company, dict) and company.get("name"):
                    company_name = company["name"]; normalized_name = normalize_company_name(company_name)
                    if not normalized_name: continue
                    try: 
                        company_id = None
                       
                        company_data = {"name": company_name, "normalized_name": normalized_name, "organization_id": organization_id}
                        supabase.table("companies").upsert(company_data, on_conflict="name, organization_id").execute()
                      
                        fetch_resp = supabase.table("companies").select("id").eq("name", company_name).eq("organization_id", organization_id).limit(1).maybe_single().execute()
                        if fetch_resp.data: company_id = fetch_resp.data["id"]
                        else: raise Exception(f"Could not retrieve company ID after upsert for: {company_name}")
                   
                        company_entries.append({"candidate_id": candidate_id, "job_id": job_id, "company_id": company_id, "designation": company.get("designation", "-"), "years": company.get("years", "-"), "organization_id": organization_id})
                    except Exception as company_proc_exc:
                        log_progress(job_id, "process_companies_error", f"Error processing company '{company_name}': {company_proc_exc}")
                        logger.warning(f"Skipping company {company_name} due to error: {company_proc_exc}")
                        continue  
            if company_entries:
                log_progress(job_id, "save_candidate_companies_start", "Upserting associations", {"count": len(company_entries)})
                try:
                    supabase.table("candidate_companies").upsert(company_entries, on_conflict="candidate_id,job_id,company_id").execute()
                    log_progress(job_id, "save_candidate_companies_success", "Upserted associations")
                except Exception as assoc_upsert_exc:
                    log_progress(job_id, "save_candidate_companies_error", f"Upsert associations failed: {assoc_upsert_exc}")
                    raise 
            else: log_progress(job_id, "save_candidate_companies_skip", "No associations to save")
        else: log_progress(job_id, "process_companies_warning", "Report 'companies' is not a list")
        log_progress(job_id, "process_companies_finished", "Finished processing companies")


        log_progress(job_id, "prepare_final_payload", "Preparing final analysis payload")
        supabase_project_id = os.getenv("SUPABASE_PROJECT_ID", "[YOUR_PROJECT_ID]")
        if supabase_project_id == "[YOUR_PROJECT_ID]":
             try: supabase_project_id = SUPABASE_URL.split('.')[0].split('//')[1]
             except Exception: log_progress(job_id, "warning", "Could not get SUPABASE_PROJECT_ID for report URL")
        report_public_url = f"https://{supabase_project_id}.supabase.co/storage/v1/object/public/{SUPABASE_STORAGE_BUCKET}/{report_destination_path}"
        resume_payload = { 
            "job_id": job_id, 
            "candidate_id": candidate_id, 
            "resume_text": resume_text or None,
            
            # --- SCORING ---
            "overall_score": round(report.get("overall_match_score", 0)),
            
            # --- TEXT/ARRAYS ---
            "summary": report.get("summary"),
            "missing_or_weak_areas": report.get("missing_or_weak_areas", []),
            "top_skills": report.get("top_skills", []), 
            "development_gaps": report.get("development_gaps", []),
            "additional_certifications": report.get("additional_certifications", []),
            
            # --- JSONB STRUCTURES (Rich Data) ---
            "matched_skills": report.get("matched_skills", []),
            "section_wise_scoring": report.get("section_wise_scoring", {}),
            "match_quality": report.get("match_quality", {}),
            "resume_quality": report.get("resume_quality", {}),
            
            # --- CANDIDATE INFO ---
            "candidate_name": report.get("candidate_name", "Unknown"), 
            "email": report.get("email", ""),
            "phone_number": report.get("phone_number", ""), 
            "github": report.get("github", ""),
            "linkedin": report.get("linkedin", ""), 
            
            # --- METADATA ---
            "report_url": report_public_url,
            "has_validated_resume": True, 
            "updated_at": datetime.datetime.utcnow().isoformat(),
            "organization_id": organization_id,
            
            # --- CRITICAL: RAW BACKUP ---
            # Saves the full object including 'experience_analysis' (stats), 
            # 'requirements_coverage', and 'red_flags' which might not have dedicated columns yet.
            "raw_ai_analysis": report
        }
        log_progress(job_id, "save_final_analysis", "Upserting final candidate_resume_analysis data")
      
        analysis_response = supabase.table("candidate_resume_analysis").upsert(
            resume_payload, on_conflict="job_id,candidate_id"
        ).execute() 
        log_progress(job_id, "save_final_analysis_response", "Final upsert response", {"data": str(analysis_response.data), "count": str(analysis_response.count)})

      
        log_progress(job_id, "success", "Task completed successfully and final data saved.")
        return {"status": "finished", "candidate_id": candidate_id}

    except Exception as e:
        
        tb_str = traceback.format_exc()
        error_message = f"Task failed during processing: {str(e)}"
        logger.error(f"Job {job_id} - process_analysis_error: {error_message} for candidate {candidate_id}")
        logger.error(f"Job {job_id} - Traceback:\n{tb_str}")
      
        log_progress(job_id, "error_processing", error_message, {"error_type": type(e).__name__, "traceback": tb_str})
       
        return {"status": "failed", "candidate_id": candidate_id, "error": str(e)}

    finally:
        
        try:
            log_progress(job_id, "cleanup", "Cleaning up temporary files")
            if local_resume_path and os.path.exists(local_resume_path): os.remove(local_resume_path)
            if local_report_path and os.path.exists(local_report_path): os.remove(local_report_path)
            log_progress(job_id, "cleanup", "Temporary files removed successfully")
        except Exception as cleanup_e:
            log_progress(job_id, "cleanup_error", f"Failed to clean up temporary files: {str(cleanup_e)}")

# --- END OF tasks.py FILE ---
# new change for the analysis_config
import os
from supabase import create_client, Client
from openai import OpenAI
from dotenv import load_dotenv
from redis import Redis
from rq import Queue

# Load environment variables from .env file (if present)
load_dotenv()

# Load environment variables
SUPABASE_URL = os.getenv("SUPABASE_URL")
SUPABASE_KEY = os.getenv("SUPABASE_KEY")
SUPABASE_SERVICE_KEY = os.environ.get("SUPABASE_SERVICE_KEY") #
SUPABASE_STORAGE_BUCKET = os.getenv("SUPABASE_STORAGE_BUCKET")
SUPABASE_RESUME_PATH_PREFIX = os.getenv("SUPABASE_RESUME_PATH_PREFIX")
SUPABASE_REPORT_PATH_PREFIX = os.getenv("SUPABASE_REPORT_PATH_PREFIX")
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
PORT = int(os.getenv("PORT", 5005))

if not all([SUPABASE_URL, SUPABASE_KEY, SUPABASE_STORAGE_BUCKET, SUPABASE_RESUME_PATH_PREFIX, SUPABASE_REPORT_PATH_PREFIX, OPENAI_API_KEY, SUPABASE_SERVICE_KEY]):
    raise ValueError("Missing required environment variables. Ensure SUPABASE_URL, SUPABASE_KEY, SUPABASE_STORAGE_BUCKET, SUPABASE_RESUME_PATH_PREFIX, SUPABASE_REPORT_PATH_PREFIX, and OPENAI_API_KEY are set.")

# Initialize Supabase client
supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY)

# !! THIS IS THE NEW ADMIN CLIENT FOR YOUR SERVER !!
supabase_admin_client: Client = create_client(SUPABASE_URL, SUPABASE_SERVICE_KEY)


# --- REMOVED: Gemini API Configuration ---
# configure(api_key=GEMINI_API_KEY)
# gemini_model = GenerativeModel("gemini-1.5-pro")

# --- ADDED: Initialize OpenAI client ---
openai_client = OpenAI(api_key=OPENAI_API_KEY)

# Initialize Redis and RQ queue
# Server host
redis_conn = Redis(host='redis', port=6379)
queue = Queue(connection=redis_conn)

# Localhost host
# redis_conn = Redis(host='localhost', port=6379)
# queue = Queue(connection=redis_conn)

# redis_conn = Redis(host='localhost', port=6379)
# queue = Queue(connection=redis_conn)
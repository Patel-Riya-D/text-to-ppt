import os
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

# =========================
# 🔹 Azure OpenAI (LLM)
# =========================
AZURE_KEY = os.getenv("AZURE_OPENAI_LLM_KEY")
AZURE_ENDPOINT = os.getenv("AZURE_LLM_ENDPOINT")
AZURE_API_VERSION = os.getenv("AZURE_LLM_API_VERSION")
AZURE_DEPLOYMENT = os.getenv("AZURE_LLM_DEPLOYMENT_41_MINI")

# =========================
# 🔹 Hugging Face (optional - for image)
# =========================
# HF_API_KEY = os.getenv("HF_API_KEY")

# =========================
# 🔹 Debug check (optional)
# =========================
if not AZURE_KEY:
    print("⚠️ Warning: AZURE_OPENAI_LLM_KEY not set")

if not AZURE_ENDPOINT:
    print("⚠️ Warning: AZURE_LLM_ENDPOINT not set")
from fastapi import FastAPI, HTTPException
from fastapi.responses import FileResponse
from pydantic import BaseModel
import os
from fastapi import File, UploadFile, Form

# Services
from backend.services.ppt_service import generate_slide_content, create_ppt
from fastapi.responses import JSONResponse
import base64

app = FastAPI(title="AI Generator API 🚀")

# =========================
# 📌 Request Model
# =========================
class PPTRequest(BaseModel):
    topic: str
    num_slides: int = 5
    tone: str = "Professional"


# =========================
# 🏠 Health Check
# =========================
@app.get("/")
def home():
    return {"message": "FastAPI backend running 🚀"}


# =========================
# 📊 Generate PPT Endpoint
# =========================
@app.post("/generate-ppt")
async def generate_ppt(
    topic: str = Form(...),
    num_slides: int = Form(5),
    tone: str = Form("Professional"),
    logo: UploadFile = File(None)
):
    try:
        # Save logo if provided
        logo_path = None
        if logo:
            os.makedirs("uploads", exist_ok=True)
            logo_path = f"uploads/{logo.filename}"

            with open(logo_path, "wb") as f:
                f.write(await logo.read())

        # Generate content
        slide_data = generate_slide_content(topic, num_slides, tone)

        file_path = create_ppt(slide_data, topic, logo_path, tone=tone)

        # Create PPT with logo
        from fastapi.responses import JSONResponse
        import base64

        # Read file
        with open(file_path, "rb") as f:
            ppt_bytes = f.read()

        # Convert to base64
        ppt_base64 = base64.b64encode(ppt_bytes).decode()

        # Return slides + file
        return JSONResponse({
            "slides": slide_data["slides"],
            "ppt_base64": ppt_base64
        })

    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
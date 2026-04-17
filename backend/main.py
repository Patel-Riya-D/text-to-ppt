# ------------------------------------------------------------------------------
#  FASTAPI BACKEND (embedded for convenience)
# ------------------------------------------------------------------------------
from fastapi import FastAPI, HTTPException
from pydantic import BaseModel
import os
import asyncio
from fastapi import File, UploadFile, Form
import json
from backend.services.ppt_service import generate_slide_content, create_ppt
from fastapi.responses import JSONResponse
import base64
import traceback

app = FastAPI(title="AI Generator API 🚀")

class PPTRequest(BaseModel):
    topic: str
    num_slides: int = 5
    tone: str = "Professional"

class PPTBuildRequest(BaseModel):
    topic: str
    tone: str = "Professional"
    slide_data: dict

@app.get("/")
def home():
    return {"message": "FastAPI backend running 🚀"}

@app.post("/generate-outline")
async def generate_outline(
    topic: str = Form(...),
    num_slides: int = Form(5),
    tone: str = Form("Professional"),
    theme_colors: str = Form(None),
):
    try:
        if theme_colors:
            theme_colors = json.loads(theme_colors)
        slide_data = await asyncio.wait_for(
            asyncio.to_thread(generate_slide_content, topic, num_slides, tone, theme_colors),
            timeout=150,
        )
        return JSONResponse(slide_data)
    except asyncio.TimeoutError:
        raise HTTPException(status_code=504, detail="Slide generation timed out")
    except Exception as e:
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/build-ppt")
async def build_ppt(
    topic: str = Form(...),
    tone: str = Form("Professional"),
    slides_json: str = Form(...),
    logo: UploadFile = File(None),
    content_image: UploadFile = File(None),
):
    try:
        slide_data = json.loads(slides_json)

        logo_path = None
        if logo:
            os.makedirs("uploads", exist_ok=True)
            logo_path = f"uploads/{logo.filename}"
            with open(logo_path, "wb") as f:
                f.write(await logo.read())

        content_image_path = None
        if content_image:
            os.makedirs("uploads", exist_ok=True)
            content_image_path = f"uploads/{content_image.filename}"
            with open(content_image_path, "wb") as f:
                f.write(await content_image.read())

        ppt_bytes = await asyncio.wait_for(
            asyncio.to_thread(create_ppt, slide_data, topic, logo_path, tone, content_image_path),
            timeout=90,
        )

        ppt_base64 = base64.b64encode(ppt_bytes).decode()

        return JSONResponse({
            "slides": slide_data.get("slides", []),
            "ppt_base64": ppt_base64,
        })
    except asyncio.TimeoutError:
        raise HTTPException(status_code=504, detail="PPT build timed out")
    except Exception as e:
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=str(e))
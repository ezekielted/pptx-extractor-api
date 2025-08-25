from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import JSONResponse

from spire.presentation import *
from spire.presentation.common import *

import os
import json
import cloudinary
import cloudinary.uploader
from dotenv import load_dotenv
import tempfile
import shutil

# --- Load Environment Variables ---
load_dotenv()

# --- Cloudinary Configuration ---
try:
    cloudinary.config(
      cloud_name = os.getenv("CLOUD_NAME"),
      api_key = os.getenv("API_KEY"),
      api_secret = os.getenv("API_SECRET"),
      secure = True
    )
    if not all([os.getenv("CLOUD_NAME"), os.getenv("API_KEY"), os.getenv("API_SECRET")]):
        raise ValueError("Cloudinary credentials not fully configured in .env")
except Exception as e:
    print(f"Error loading Cloudinary configuration: {e}")


app = FastAPI(
    title="PowerPoint Extractor API",
    description="API to extract text and images from PPTX files using Spire.Presentation and Cloudinary.",
    version="1.0.0"
)

@app.post("/extract-pptx")
async def extract_pptx(
    File: UploadFile = File(..., description="The PowerPoint presentation file (.pptx)"),
    extractText: bool = Form(False, description="Set to true to extract text content."),
    extractImage: bool = Form(False, description="Set to true to extract images (charts, pictures) and upload to Cloudinary."),
    extractAll: bool = Form(False, description="Set to true to extract both text and images. This flag overrides `extractText` and `extractImage` if set to true.")
):
    if not File.filename.endswith(('.pptx', '.ppt')):
        raise HTTPException(
            status_code=400,
            detail="Invalid file type. Only .pptx or .ppt files are accepted."
        )

    if extractAll:
        extractText = True
        extractImage = True

    if not extractText and not extractImage:
        raise HTTPException(
            status_code=400,
            detail="At least one extraction option (extractText, extractImage, or extractAll) must be true."
        )

    ppt = None
    presentation_data = {"slides": []}
    temp_pptx_path = None

    try:
        ppt = Presentation()
        
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as temp_file:
            contents = await File.read()
            temp_file.write(contents)
            temp_pptx_path = temp_file.name

        ppt.LoadFromFile(temp_pptx_path)

        # --- FIX: Efficiently extract all images from the presentation ---
        if extractImage:
            image_urls = []
            for i, image in enumerate(ppt.Images):
                temp_image_file = None
                try:
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tf:
                        temp_image_file = tf.name
                    
                    # Save the embedded image to a temporary file
                    image.Image.Save(temp_image_file)
                    
                    # Upload to Cloudinary
                    upload_result = cloudinary.uploader.upload(
                        temp_image_file,
                        folder=f"pptx_extractions/{os.path.basename(File.filename).split('.')[0]}/all_images"
                    )
                    image_urls.append(upload_result['secure_url'])
                    print(f"Uploaded image {i + 1} to Cloudinary.")
                
                except Exception as e:
                    print(f"Error processing image {i + 1}: {e}")
                    image_urls.append({"error": f"Failed to extract/upload image {i + 1}"})
                
                finally:
                    if temp_image_file and os.path.exists(temp_image_file):
                        os.remove(temp_image_file)
            
            presentation_data["extracted_images"] = image_urls

        for slide_index, slide in enumerate(ppt.Slides):
            slide_content = { "slide": slide_index + 1 }

            if extractText:
                slide_content["text"] = []
                for shape in slide.Shapes:
                    if isinstance(shape, IAutoShape) and shape.TextFrame is not None:
                        for paragraph in shape.TextFrame.Paragraphs:
                            if paragraph.Text.strip():
                                slide_content["text"].append(paragraph.Text)
            
            presentation_data["slides"].append(slide_content)

        return JSONResponse(content=presentation_data)

    except Exception as e:
        raise HTTPException(
            status_code=500,
            detail=f"An error occurred during PPTX processing: {e}"
        )
    finally:
        if ppt is not None:
            ppt.Dispose()
        if temp_pptx_path and os.path.exists(temp_pptx_path):
            os.remove(temp_pptx_path)
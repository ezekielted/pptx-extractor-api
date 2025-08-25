from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import JSONResponse

# Make sure to import all necessary classes
from spire.presentation import Presentation, IChart, PictureShape, IAutoShape

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

        for slide_index, slide in enumerate(ppt.Slides):
            slide_content = { "slide": slide_index + 1 }

            # --- Text extraction logic (unchanged) ---
            if extractText:
                slide_content["text"] = []
                for shape in slide.Shapes:
                    if isinstance(shape, IAutoShape) and shape.TextFrame is not None:
                        for paragraph in shape.TextFrame.Paragraphs:
                            if paragraph.Text.strip():
                                slide_content["text"].append(paragraph.Text)

            # --- Image and Chart extraction logic (fixed and integrated) ---
            if extractImage:
                slide_content["images"] = []

                for shape in slide.Shapes:
                    image_to_save = None
                    temp_image_file = None
                    
                    try:
                        # --- Logic for Charts (preserved from your original code) ---
                        if isinstance(shape, IChart):
                            image_to_save = shape.SaveAsImage()

                        # --- CORRECTED LOGIC: Added for actual Images ---
                        elif isinstance(shape, PictureShape):
                            if shape.Picture is not None and shape.Picture.Image is not None:
                                image_to_save = shape.Picture.Image
                        
                        # If a chart or image was found, save and upload it
                        if image_to_save is not None:
                            with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tf:
                                temp_image_file = tf.name
                            
                            # Save the image data to the temporary file
                            image_to_save.Save(temp_image_file)
                        
                            # Upload to Cloudinary
                            upload_result = cloudinary.uploader.upload(
                                temp_image_file, 
                                folder=f"pptx_extractions/{os.path.basename(File.filename).split('.')[0]}/slide_{slide_index + 1}"
                            )
                            slide_content["images"].append(upload_result['secure_url'])
                            print(f"Uploaded image/chart from slide {slide_index + 1} to Cloudinary.")
                        
                    except Exception as e:
                        print(f"Error processing a shape on slide {slide_index + 1}: {e}")
                        if "images" not in slide_content:
                            slide_content["images"] = []
                        slide_content["images"].append({"error": "Failed to extract/upload an image or chart"})
                    finally:
                        # Clean up the temporary image file
                        if temp_image_file and os.path.exists(temp_image_file):
                            os.remove(temp_image_file)
            
            # Only add the slide to the results if it has content
            if slide_content.get("text") or slide_content.get("images"):
                presentation_data["slides"].append(slide_content)

        return JSONResponse(content=presentation_data)

    except Exception as e:
        raise HTTPException(
            status_code=500,
            detail=f"An error occurred during PPTX processing: {str(e)}"
        )
    finally:
        if ppt is not None:
            ppt.Dispose()
        if temp_pptx_path and os.path.exists(temp_pptx_path):
            os.remove(temp_pptx_path)
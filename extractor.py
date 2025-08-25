from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import JSONResponse

from spire.presentation import *

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

        # Extract all images from the presentation first (if needed)
        extracted_images = []
        if extractImage:
            print(f"Found {len(ppt.Images)} images in the presentation")
            for i, image in enumerate(ppt.Images):
                temp_image_file = None
                try:
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tf:
                        temp_image_file = tf.name
                    
                    # Save the image to temporary file
                    image.Image.Save(temp_image_file)
                    
                    # Upload to Cloudinary
                    upload_result = cloudinary.uploader.upload(
                        temp_image_file, 
                        folder=f"pptx_extractions/{os.path.basename(File.filename).split('.')[0]}/images"
                    )
                    
                    extracted_images.append({
                        "index": i,
                        "url": upload_result['secure_url'],
                        "public_id": upload_result['public_id']
                    })
                    
                    print(f"Uploaded image {i} to Cloudinary: {upload_result['secure_url']}")
                    
                except Exception as e:
                    print(f"Error processing image {i}: {e}")
                    extracted_images.append({
                        "index": i,
                        "error": f"Failed to extract/upload image {i}: {str(e)}"
                    })
                finally:
                    if temp_image_file and os.path.exists(temp_image_file):
                        os.remove(temp_image_file)

        # Process each slide
        for slide_index, slide in enumerate(ppt.Slides):
            slide_content = {"slide": slide_index + 1}

            if extractText:
                slide_content["text"] = []
                for shape in slide.Shapes:
                    if isinstance(shape, IAutoShape) and shape.TextFrame is not None:
                        for paragraph in shape.TextFrame.Paragraphs:
                            if paragraph.Text.strip():
                                slide_content["text"].append(paragraph.Text)

            if extractImage:
                slide_content["images"] = []
                image_count_on_slide = 1

                # Process charts (keeping your existing chart extraction logic)
                for shape in slide.Shapes:
                    temp_image_file = None
                    try:
                        image_to_save = None
                        
                        if isinstance(shape, IChart):
                            image_to_save = shape.SaveAsImage()
                            
                            if image_to_save is not None:
                                with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tf:
                                    temp_image_file = tf.name
                                image_to_save.Save(temp_image_file)
                            
                                upload_result = cloudinary.uploader.upload(
                                    temp_image_file, 
                                    folder=f"pptx_extractions/{os.path.basename(File.filename).split('.')[0]}/slide_{slide_index + 1}/charts"
                                )
                                slide_content["images"].append({
                                    "type": "chart",
                                    "url": upload_result['secure_url'],
                                    "public_id": upload_result['public_id']
                                })
                                print(f"Uploaded chart from slide {slide_index + 1}, shape {image_count_on_slide} to Cloudinary.")
                        
                    except Exception as e:
                        print(f"Error processing chart from slide {slide_index + 1}, shape {image_count_on_slide}: {e}")
                        if "images" not in slide_content:
                            slide_content["images"] = []
                        slide_content["images"].append({
                            "type": "chart",
                            "error": f"Failed to extract/upload chart for shape {image_count_on_slide}: {str(e)}"
                        })
                    finally:
                        if temp_image_file and os.path.exists(temp_image_file):
                            os.remove(temp_image_file)
                    image_count_on_slide += 1

                # Add all extracted images to each slide (you might want to modify this logic
                # to only include images that are actually on this specific slide)
                for img_data in extracted_images:
                    if "error" not in img_data:
                        slide_content["images"].append({
                            "type": "image",
                            "url": img_data["url"],
                            "public_id": img_data["public_id"],
                            "image_index": img_data["index"]
                        })
                    else:
                        slide_content["images"].append({
                            "type": "image",
                            "error": img_data["error"],
                            "image_index": img_data["index"]
                        })
            
            presentation_data["slides"].append(slide_content)

        # Also add a summary of all extracted images at the presentation level
        if extractImage and extracted_images:
            presentation_data["all_images"] = extracted_images

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
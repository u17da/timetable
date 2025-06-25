from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse
import os
import base64
from io import BytesIO
from PIL import Image
import openpyxl
from openai import OpenAI
from dotenv import load_dotenv
import json
from typing import Dict, Any, List

load_dotenv()
print(f"DEBUG: OPENAI_API_KEY from environment: {os.getenv('OPENAI_API_KEY')[:20]}...{os.getenv('OPENAI_API_KEY')[-10:] if os.getenv('OPENAI_API_KEY') else 'None'}")

app = FastAPI()

# Disable CORS. Do not remove this for full-stack development.
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # Allows all origins
    allow_credentials=True,
    allow_methods=["*"],  # Allows all methods
    allow_headers=["*"],  # Allows all headers
)

openai_api_key = os.getenv("OPENAI_API_KEY")
print(f"DEBUG: OpenAI client initialization with key: {openai_api_key[:20]}...{openai_api_key[-10:] if openai_api_key else 'None'}")
openai_client = OpenAI(api_key=openai_api_key)

timetables_storage: Dict[str, Any] = {}

@app.get("/healthz")
async def healthz():
    return {"status": "ok"}

@app.post("/upload")
async def upload_file(file: UploadFile = File(...)):
    """Upload and parse a timetable file (image or Excel)"""
    try:
        print(f"DEBUG: Received file upload - filename: {file.filename}, content_type: {file.content_type}, size: {file.size}")
        
        if file.content_type.startswith('image/'):
            print("DEBUG: Processing as image file")
            return await process_image_file(file)
        elif file.content_type in ['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'application/vnd.ms-excel']:
            print("DEBUG: Processing as Excel file")
            return await process_excel_file(file)
        else:
            print(f"DEBUG: Unsupported file type: {file.content_type}")
            raise HTTPException(status_code=400, detail="Unsupported file type. Please upload an image or Excel file.")
    
    except Exception as e:
        print(f"DEBUG: Exception in upload_file: {type(e).__name__}: {str(e)}")
        raise HTTPException(status_code=500, detail=f"Error processing file: {str(e)}")

async def process_image_file(file: UploadFile) -> JSONResponse:
    """Process uploaded image file using OpenAI Vision API"""
    try:
        image_data = await file.read()
        image = Image.open(BytesIO(image_data))
        
        buffered = BytesIO()
        image.save(buffered, format="PNG")
        img_base64 = base64.b64encode(buffered.getvalue()).decode()
        
        print(f"DEBUG: About to call OpenAI API with client using key: {openai_client.api_key[:20]}...{openai_client.api_key[-10:] if openai_client.api_key else 'None'}")
        response = openai_client.chat.completions.create(
            model="gpt-4o",
            messages=[
                {
                    "role": "user",
                    "content": [
                        {
                            "type": "text",
                            "text": """Please analyze this timetable image and extract the schedule information. Return a JSON object with the following structure:
                            {
                                "title": "Schedule title if visible",
                                "schedule": {
                                    "Monday": [{"time": "09:00-10:00", "subject": "Math", "room": "A101"}],
                                    "Tuesday": [{"time": "09:00-10:00", "subject": "English", "room": "B202"}],
                                    "Wednesday": [],
                                    "Thursday": [],
                                    "Friday": [],
                                    "Saturday": [],
                                    "Sunday": []
                                }
                            }
                            Extract all visible time slots, subjects, and room numbers. If information is unclear, use your best judgment."""
                        },
                        {
                            "type": "image_url",
                            "image_url": {
                                "url": f"data:image/png;base64,{img_base64}"
                            }
                        }
                    ]
                }
            ],
            max_tokens=1000
        )
        
        content = response.choices[0].message.content
        
        try:
            timetable_data = json.loads(content)
        except json.JSONDecodeError:
            import re
            json_match = re.search(r'```(?:json)?\s*(\{.*?\})\s*```', content, re.DOTALL)
            if json_match:
                timetable_data = json.loads(json_match.group(1))
            else:
                json_match = re.search(r'\{.*\}', content, re.DOTALL)
                if json_match:
                    timetable_data = json.loads(json_match.group(0))
                else:
                    raise json.JSONDecodeError("No valid JSON found in OpenAI response", content, 0)
        
        file_id = f"img_{len(timetables_storage)}"
        timetables_storage[file_id] = timetable_data
        
        return JSONResponse(content={"id": file_id, "data": timetable_data})
        
    except json.JSONDecodeError:
        raise HTTPException(status_code=500, detail="Failed to parse OpenAI response as JSON")
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error processing image: {str(e)}")

async def process_excel_file(file: UploadFile) -> JSONResponse:
    """Process uploaded Excel file using OpenAI to structure the data"""
    try:
        excel_data = await file.read()
        workbook = openpyxl.load_workbook(BytesIO(excel_data))
        sheet = workbook.active
        
        excel_content = []
        for row in sheet.iter_rows(values_only=True):
            if any(cell is not None for cell in row):
                excel_content.append([str(cell) if cell is not None else "" for cell in row])
        
        excel_text = "\n".join(["\t".join(row) for row in excel_content])
        
        response = openai_client.chat.completions.create(
            model="gpt-4o",
            messages=[
                {
                    "role": "user",
                    "content": f"""Please analyze this Excel timetable data and convert it to a structured JSON format. The data is:

{excel_text}

Return a JSON object with the following structure:
{{
    "title": "Schedule title if identifiable",
    "schedule": {{
        "Monday": [{{"time": "09:00-10:00", "subject": "Math", "room": "A101"}}],
        "Tuesday": [{{"time": "09:00-10:00", "subject": "English", "room": "B202"}}],
        "Wednesday": [],
        "Thursday": [],
        "Friday": [],
        "Saturday": [],
        "Sunday": []
    }}
}}

Extract all time slots, subjects, and room information. Organize by weekdays. If the format is unclear, use your best judgment to structure the data appropriately."""
                }
            ],
            max_tokens=1000
        )
        
        content = response.choices[0].message.content
        
        try:
            timetable_data = json.loads(content)
        except json.JSONDecodeError:
            import re
            json_match = re.search(r'```(?:json)?\s*(\{.*?\})\s*```', content, re.DOTALL)
            if json_match:
                timetable_data = json.loads(json_match.group(1))
            else:
                json_match = re.search(r'\{.*\}', content, re.DOTALL)
                if json_match:
                    timetable_data = json.loads(json_match.group(0))
                else:
                    raise json.JSONDecodeError("No valid JSON found in OpenAI response", content, 0)
        
        file_id = f"excel_{len(timetables_storage)}"
        timetables_storage[file_id] = timetable_data
        
        return JSONResponse(content={"id": file_id, "data": timetable_data})
        
    except json.JSONDecodeError:
        raise HTTPException(status_code=500, detail="Failed to parse OpenAI response as JSON")
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error processing Excel file: {str(e)}")

@app.get("/timetable/{file_id}")
async def get_timetable(file_id: str):
    """Retrieve a stored timetable by ID"""
    if file_id not in timetables_storage:
        raise HTTPException(status_code=404, detail="Timetable not found")
    
    return JSONResponse(content=timetables_storage[file_id])

@app.get("/timetables")
async def list_timetables():
    """List all stored timetables"""
    return JSONResponse(content={
        "timetables": [
            {"id": file_id, "title": data.get("title", "Untitled")}
            for file_id, data in timetables_storage.items()
        ]
    })

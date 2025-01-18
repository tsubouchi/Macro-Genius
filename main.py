import os
from fastapi import FastAPI, HTTPException, Request
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from pydantic import BaseModel
from typing import List
import openpyxl
from datetime import datetime
import database as db

app = FastAPI()

app.mount("/static", StaticFiles(directory="static"), name="static")
templates = Jinja2Templates(directory="templates")

class MacroRequest(BaseModel):
    description: str

class MacroResponse(BaseModel):
    id: int
    description: str
    created_at: str

@app.get("/")
async def read_root(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})

@app.post("/generate-macro")
async def generate_macro(macro_request: MacroRequest):
    try:
        # 仮のマクロ生成ロジック
        wb = openpyxl.Workbook()
        ws = wb.active
        ws['A1'] = 'Generated Macro'
        ws['A2'] = macro_request.description
        
        # ファイル保存
        filename = f"macro_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        filepath = f"temp/{filename}"
        os.makedirs("temp", exist_ok=True)
        wb.save(filepath)
        
        # DBに保存
        macro_id = db.save_macro(macro_request.description)
        
        return FileResponse(
            filepath,
            filename=filename,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/macros", response_model=List[MacroResponse])
async def get_macros():
    try:
        macros = db.get_all_macros()
        return macros
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)

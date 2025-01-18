import os
from fastapi import FastAPI, HTTPException, Request
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from typing import List
import openpyxl
from datetime import datetime
import database as db
from models import Macro, MacroRequest, MacroCategory # Added import
import openai # Added import

app = FastAPI()

app.mount("/static", StaticFiles(directory="static"), name="static")
templates = Jinja2Templates(directory="templates")

openai.api_key = os.getenv("OPENAI_API_KEY")

@app.get("/")
async def read_root(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})

@app.post("/generate-macro")
async def generate_macro(macro_request: MacroRequest):
    try:
        if macro_request.template_id:
            # テンプレートからマクロを生成
            template = db.get_macro_by_id(macro_request.template_id)
            if not template:
                raise HTTPException(status_code=404, detail="Template not found")
            description = template["description"]
        elif macro_request.use_ai:
            # GPT-4を使用してマクロの詳細を生成
            if not macro_request.description:
                raise HTTPException(status_code=400, detail="Description is required for AI generation")

            response = openai.ChatCompletion.create(
                model="gpt-4",
                messages=[
                    {"role": "system", "content": "あなたはExcelマクロの専門家です。ユーザーの要件に基づいて、実用的なVBAマクロを生成してください。"},
                    {"role": "user", "content": f"以下の要件に基づくExcelマクロを作成してください：\n{macro_request.description}"}
                ]
            )
            description = response.choices[0].message.content
        else:
            raise HTTPException(status_code=400, detail="Either template_id or description with use_ai must be provided")

        # Excelファイル生成
        wb = openpyxl.Workbook()
        ws = wb.active
        ws['A1'] = 'Generated Macro'
        ws['A2'] = description

        filename = f"macro_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        filepath = f"temp/{filename}"
        os.makedirs("temp", exist_ok=True)
        wb.save(filepath)

        # DBに保存
        macro_id = db.save_macro(
            title=f"マクロ {datetime.now().strftime('%Y/%m/%d %H:%M')}",
            description=description,
            category=MacroCategory.AI_GENERATED if macro_request.use_ai else MacroCategory.TEMPLATE
        )

        return FileResponse(
            filepath,
            filename=filename,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/macros", response_model=List[Macro]) # Updated response model
async def get_macros():
    try:
        macros = db.get_all_macros()
        return macros
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
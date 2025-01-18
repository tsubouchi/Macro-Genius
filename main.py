import os
from fastapi import FastAPI, HTTPException, Request
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from typing import List
import openpyxl
from datetime import datetime
import database as db
from models import Macro, MacroRequest, MacroCategory
import openai

# FastAPIアプリケーションの設定
app = FastAPI()

# 静的ファイルとテンプレートの設定
app.mount("/static", StaticFiles(directory="static"), name="static")
templates = Jinja2Templates(directory="templates")

# OpenAI APIキーの設定
openai.api_key = os.getenv("OPENAI_API_KEY")

@app.get("/")
async def read_root(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})

@app.post("/generate-macro")
async def generate_macro(macro_request: MacroRequest):
    try:
        if macro_request.template_id:
            template = db.get_macro_by_id(macro_request.template_id)
            if not template:
                raise HTTPException(status_code=404, detail="Template not found")

            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "マクロ"
            ws['A1'] = template["title"]
            ws['A2'] = template["description"]

            filename = f"macro_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            filepath = f"temp/{filename}"
            os.makedirs("temp", exist_ok=True)
            wb.save(filepath)

            return FileResponse(
                filepath,
                filename=filename,
                media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        elif macro_request.use_ai:
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

            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "マクロ"
            ws['A1'] = f"AI生成マクロ {datetime.now().strftime('%Y/%m/%d %H:%M')}"
            ws['A2'] = description

            filename = f"macro_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            filepath = f"temp/{filename}"
            os.makedirs("temp", exist_ok=True)
            wb.save(filepath)

            macro_id = db.save_macro(
                title=f"マクロ {datetime.now().strftime('%Y/%m/%d %H:%M')}",
                description=description,
                category=MacroCategory.AI_GENERATED
            )

            return FileResponse(
                filepath,
                filename=filename,
                media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            raise HTTPException(status_code=400, detail="Either template_id or description with use_ai must be provided")
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/macros", response_model=List[Macro])
async def get_macros():
    try:
        macros = db.get_all_macros()
        return macros
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

# テンプレートマクロの初期データを作成
@app.on_event("startup")
async def startup_event():
    try:
        # テンプレートマクロの作成
        templates = [
            {
                "title": "データ集計マクロ",
                "description": """
Sub データ集計()
    ' 選択範囲の合計を計算
    Dim rng As Range
    Set rng = Selection

    ' 合計を計算
    Dim total As Double
    total = WorksheetFunction.Sum(rng)

    ' 結果を表示
    MsgBox "選択範囲の合計: " & total
End Sub
"""
            },
            {
                "title": "シート整理マクロ",
                "description": """
Sub シート整理()
    ' すべてのシートをループ
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ' シートの最終行と列を取得
        Dim lastRow As Long, lastCol As Long
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

        ' オートフィット
        ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).Columns.AutoFit
    Next ws

    MsgBox "すべてのシートを整理しました"
End Sub
"""
            },
            {
                "title": "データ検索マクロ",
                "description": """
Sub データ検索()
    ' 検索値の入力
    Dim searchValue As String
    searchValue = InputBox("検索する値を入力してください")

    ' 検索実行
    Dim foundCell As Range
    Set foundCell = ActiveSheet.Range("A:Z").Find(searchValue)

    ' 結果の処理
    If Not foundCell Is Nothing Then
        foundCell.Select
        MsgBox "セル " & foundCell.Address & " に見つかりました"
    Else
        MsgBox "見つかりませんでした"
    End If
End Sub
"""
            }
        ]

        for template in templates:
            db.save_macro(
                title=template["title"],
                description=template["description"],
                category=MacroCategory.TEMPLATE
            )

    except Exception as e:
        print(f"Error creating templates: {str(e)}")

# ローカル開発用の設定を更新
if __name__ == "__main__":
    import uvicorn
    port = int(os.getenv('PORT', 3000))
    uvicorn.run(app, host="0.0.0.0", port=port)
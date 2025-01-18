import os
from flask import Flask
from database import db

app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_SECRET_KEY") or "a secret key"

# Configure database
app.config["SQLALCHEMY_DATABASE_URI"] = os.environ.get("DATABASE_URL")
app.config["SQLALCHEMY_ENGINE_OPTIONS"] = {
    "pool_recycle": 300,
    "pool_pre_ping": True,
}

# Initialize database
db.init_app(app)

def init_templates():
    """Initialize template macros if they don't exist"""
    from models import Macro, MacroCategory

    # Check if templates already exist
    templates_exist = db.session.query(db.exists().where(
        Macro.category == MacroCategory.TEMPLATE
    )).scalar()

    if not templates_exist:
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
            }
        ]

        for template in templates:
            macro = Macro(
                title=template["title"],
                description=template["description"],
                category=MacroCategory.TEMPLATE,
                is_public=True
            )
            db.session.add(macro)
            macro.add_version(template["description"])

        db.session.commit()

# Initialize database tables and templates
with app.app_context():
    db.create_all()
    init_templates()
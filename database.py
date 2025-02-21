from flask_sqlalchemy import SQLAlchemy
from sqlalchemy.orm import DeclarativeBase

class Base(DeclarativeBase):
    pass

db = SQLAlchemy(model_class=Base)

def init_db():
    """データベーステーブルの初期化"""
    from models import Macro, MacroVersion  # 循環参照を避けるためにここでインポート
    db.create_all()

def save_macro(title: str, description: str, category: str = 'AI_GENERATED', is_public: bool = True) -> int:
    """新しいマクロを保存"""
    from models import Macro  # 循環参照を避けるためにここでインポート
    macro = Macro(
        title=title,
        description=description,
        category=category,
        is_public=is_public
    )
    db.session.add(macro)
    db.session.commit()
    return macro.id

def get_macro_by_id(macro_id: int):
    """IDによるマクロの取得"""
    from models import Macro
    return Macro.query.get(macro_id)

def get_all_macros(public_only: bool = False):
    """全マクロの取得"""
    from models import Macro
    query = Macro.query
    if public_only:
        query = query.filter_by(is_public=True)
    return query.order_by(Macro.created_at.desc()).all()

# DB初期化
# Note:  This needs to be called after the app is configured with the database URI
# Example in app.py:
# app.config['SQLALCHEMY_DATABASE_URI'] = os.environ.get("DATABASE_URL")
# db.init_app(app)
# with app.app_context():
#     init_db()
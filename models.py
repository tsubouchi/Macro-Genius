from datetime import datetime
from enum import Enum
from database import db

class MacroCategory(str, Enum):
    TEMPLATE = "TEMPLATE"
    AI_GENERATED = "AI_GENERATED"
    DATA_PROCESSING = "DATA_PROCESSING"
    FORMATTING = "FORMATTING"
    CALCULATION = "CALCULATION"
    AUTOMATION = "AUTOMATION"
    REPORTING = "REPORTING"
    CUSTOM = "CUSTOM"

    @classmethod
    def get_japanese_name(cls, category):
        mapping = {
            cls.TEMPLATE: "テンプレート",
            cls.AI_GENERATED: "AI生成",
            cls.DATA_PROCESSING: "データ処理",
            cls.FORMATTING: "フォーマット",
            cls.CALCULATION: "計算",
            cls.AUTOMATION: "自動化",
            cls.REPORTING: "レポート",
            cls.CUSTOM: "カスタム"
        }
        return mapping.get(category, category)

class MacroVersion(db.Model):
    __tablename__ = 'macro_versions'

    id = db.Column(db.Integer, primary_key=True)
    macro_id = db.Column(db.Integer, db.ForeignKey('macros.id', ondelete='CASCADE'), nullable=False)
    version_number = db.Column(db.Integer, nullable=False)
    content = db.Column(db.Text, nullable=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    def to_dict(self):
        return {
            'id': self.id,
            'macro_id': self.macro_id,
            'version_number': self.version_number,
            'content': self.content,
            'created_at': self.created_at.isoformat()
        }

class Macro(db.Model):
    __tablename__ = 'macros'

    id = db.Column(db.Integer, primary_key=True)
    title = db.Column(db.String(255))
    description = db.Column(db.Text, nullable=False)
    category = db.Column(db.String(50), default=MacroCategory.AI_GENERATED)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    is_public = db.Column(db.Boolean, default=False)
    versions = db.relationship('MacroVersion', backref='macro', lazy=True, cascade='all, delete-orphan')

    def get_latest_version(self):
        from sqlalchemy import desc
        return (MacroVersion.query
                .filter_by(macro_id=self.id)
                .order_by(desc(MacroVersion.version_number))
                .first())

    def add_version(self, content):
        latest = self.get_latest_version()
        next_version = 1 if not latest else latest.version_number + 1
        version = MacroVersion(
            macro_id=self.id,
            version_number=next_version,
            content=content
        )
        db.session.add(version)
        db.session.commit()
        return version

    def to_dict(self):
        latest_version = self.get_latest_version()
        return {
            'id': self.id,
            'title': self.title,
            'description': self.description,
            'category': self.category,
            'created_at': self.created_at.isoformat(),
            'is_public': self.is_public,
            'latest_version': latest_version.version_number if latest_version else None,
            'content': latest_version.content if latest_version else None
        }
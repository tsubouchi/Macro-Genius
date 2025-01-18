from datetime import datetime
from pydantic import BaseModel
from enum import Enum

class MacroCategory(str, Enum):
    TEMPLATE = "TEMPLATE"
    AI_GENERATED = "AI_GENERATED"

class Macro(BaseModel):
    id: int
    title: str | None = None
    description: str
    category: MacroCategory = MacroCategory.AI_GENERATED
    created_at: datetime

class MacroRequest(BaseModel):
    description: str
    use_ai: bool = True
    template_id: int | None = None
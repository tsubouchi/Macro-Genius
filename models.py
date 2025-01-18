from datetime import datetime
from pydantic import BaseModel

class Macro(BaseModel):
    id: int
    description: str
    created_at: datetime

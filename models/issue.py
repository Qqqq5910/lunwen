from pydantic import BaseModel
from typing import Optional, Dict, Any


class Issue(BaseModel):
    type: str
    label: Optional[str] = None
    group: str = "manual"
    paragraph_index: Optional[int] = None
    text: Optional[str] = None
    problem: str
    suggestion: str
    auto_fixable: bool = False
    fixed: bool = False
    meta: Optional[Dict[str, Any]] = None

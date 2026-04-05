import os
from google import genai
from pydantic import BaseModel

class SlideContent(BaseModel):
    title: str
    punchline: str
    layout_type: str
    bullet_points: list[str]
    icon_keyword: str

try:
    client = genai.Client()
    print("Client initialized")
except Exception as e:
    print(f"Failed: {e}")

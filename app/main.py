"""
MailMind Hosted Backend — Render + Anthropic Claude
No RAG dependency — Claude handles drafting directly
"""

from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel
from pydantic_settings import BaseSettings
import httpx
import logging
import os

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("mailmind")

# ─── SETTINGS ────────────────────────────────────────────────────────────────

class Settings(BaseSettings):
    ANTHROPIC_API_KEY: str = ""
    ANTHROPIC_MODEL: str = "claude-haiku-4-5-20251001"

    class Config:
        env_file = ".env"

settings = Settings()

# ─── APP ─────────────────────────────────────────────────────────────────────

app = FastAPI(title="MailMind API", version="2.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

# Serve Outlook add-in files if folder exists
if os.path.exists("outlook-addin"):
    app.mount("/outlook", StaticFiles(directory="outlook-addin"), name="outlook")

# ─── MODELS ──────────────────────────────────────────────────────────────────

class EmailData(BaseModel):
    id: str = ""
    threadId: str = ""
    conversationId: str = ""
    subject: str = ""
    from_: str = ""
    fromName: str = ""
    body: str = ""
    snippet: str = ""

    class Config:
        populate_by_name = True
        fields = {'from_': 'from'}

class DraftResponse(BaseModel):
    draft: str
    engine: str = "claude"

# ─── ROUTES ──────────────────────────────────────────────────────────────────

@app.get("/")
async def root():
    return {"name": "MailMind API", "version": "2.0.0", "status": "running"}


@app.get("/health")
async def health():
    return {
        "status": "ok",
        "engine": "claude",
        "model": settings.ANTHROPIC_MODEL,
        "api_key_set": bool(settings.ANTHROPIC_API_KEY)
    }


@app.post("/draft", response_model=DraftResponse)
async def draft_reply(email: EmailData):
    logger.info(f"Drafting reply for: {email.subject[:60]}")

    body = email.body or email.snippet or ""
    if not body.strip():
        raise HTTPException(status_code=400, detail="Email body is empty")

    if not settings.ANTHROPIC_API_KEY:
        raise HTTPException(status_code=500, detail="ANTHROPIC_API_KEY not set in environment")

    prompt = build_prompt(email)

    try:
        draft = await call_claude(prompt)
        return DraftResponse(draft=draft.strip(), engine="claude")
    except Exception as e:
        logger.error(f"Claude error: {e}")
        raise HTTPException(status_code=500, detail=str(e))


# ─── CLAUDE ──────────────────────────────────────────────────────────────────

async def call_claude(prompt: str) -> str:
    async with httpx.AsyncClient(timeout=30.0) as client:
        res = await client.post(
            "https://api.anthropic.com/v1/messages",
            headers={
                "x-api-key": settings.ANTHROPIC_API_KEY,
                "anthropic-version": "2023-06-01",
                "content-type": "application/json"
            },
            json={
                "model": settings.ANTHROPIC_MODEL,
                "max_tokens": 512,
                "system": (
                    "You are a personal email assistant. "
                    "Write concise, natural email replies on behalf of the user. "
                    "Never add a subject line. Get straight to the point. "
                    "Match the tone of the original email — casual if casual, formal if formal."
                ),
                "messages": [{"role": "user", "content": prompt}]
            }
        )
        res.raise_for_status()
        return res.json()["content"][0]["text"]


# ─── PROMPT ──────────────────────────────────────────────────────────────────

def build_prompt(email: EmailData) -> str:
    sender = email.fromName or email.from_ or "the sender"
    subject = email.subject or "this email"
    body = (email.body or email.snippet or "")[:1500]

    return f"""Draft a reply to this email on my behalf.

EMAIL:
From: {sender}
Subject: {subject}

{body}

---
Write only the reply body. No subject line. Be concise and natural.
Reply:"""

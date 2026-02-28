"""
MailMind Hosted Backend — Render + Anthropic Claude
Serves both Gmail extension and Outlook Add-in
"""

from fastapi import FastAPI, HTTPException, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse
from pydantic import BaseModel
from typing import Optional
import logging
import httpx
import os

from .config import settings
from .rag import RAGEngine

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("mailmind")

app = FastAPI(title="MailMind API", version="2.0.0")

# CORS — allow Outlook, Gmail, and any browser-based add-in
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

rag_engine = RAGEngine()

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
    rag_used: bool = False

class IngestRequest(BaseModel):
    emails: list[dict]

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
        "rag_docs": rag_engine.count()
    }


@app.post("/draft", response_model=DraftResponse)
async def draft_reply(email: EmailData):
    logger.info(f"Drafting reply for: {email.subject[:60]}")

    if not email.body and not email.snippet:
        raise HTTPException(status_code=400, detail="Email body is empty")

    # RAG context
    rag_context = rag_engine.retrieve(
        query=f"{email.subject} {(email.body or email.snippet)[:200]}",
        n_results=settings.MAX_RAG_RESULTS
    )
    style = rag_engine.get_style_profile()

    # Build prompt
    prompt = build_prompt(email, rag_context, style)

    # Generate with Claude
    try:
        draft = await generate_with_claude(prompt)
        return DraftResponse(
            draft=draft.strip(),
            engine="claude",
            rag_used=len(rag_context) > 0
        )
    except Exception as e:
        logger.error(f"Claude generation failed: {e}")
        raise HTTPException(status_code=500, detail=f"AI generation failed: {str(e)}")


@app.post("/ingest")
async def ingest_emails(request: IngestRequest):
    logger.info(f"Ingesting {len(request.emails)} emails")
    count = rag_engine.ingest(request.emails)
    profile = rag_engine.build_style_profile(request.emails)
    return {
        "ingested": count,
        "style_profile": profile,
        "message": f"Successfully indexed {count} emails"
    }


@app.get("/style-profile")
async def get_style_profile():
    return rag_engine.get_style_profile()


# ─── CLAUDE API ──────────────────────────────────────────────────────────────

async def generate_with_claude(prompt: str) -> str:
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
                "max_tokens": 1024,
                "system": "You are a personal email assistant. Write concise, natural email replies that match the user's writing style. Never add a subject line. Never use overly formal language unless the examples show it. Get straight to the point.",
                "messages": [{"role": "user", "content": prompt}]
            }
        )
        res.raise_for_status()
        data = res.json()
        return data["content"][0]["text"]


# ─── PROMPT BUILDER ──────────────────────────────────────────────────────────

def build_prompt(email: EmailData, rag_context: list, style: dict) -> str:
    sender = email.fromName or email.from_ or "the sender"
    subject = email.subject or "this email"
    body = (email.body or email.snippet or "")[:1200]

    examples_section = ""
    if rag_context:
        examples = "\n---\n".join([
            f"Email received: {c.get('received', '')[:200]}\nMy reply: {c.get('reply', '')[:300]}"
            for c in rag_context[:3]
        ])
        examples_section = f"\nExamples of how I reply to similar emails:\n{examples}\n"

    style_section = ""
    if style:
        style_section = f"""
My writing style:
- Tone: {style.get('tone', 'professional')}
- Length: {style.get('avg_length', 'concise')}
- Sign-off: {style.get('sign_off', 'Best regards')}
"""

    return f"""Draft a reply to this email on my behalf.
{style_section}
{examples_section}
EMAIL TO REPLY TO:
From: {sender}
Subject: {subject}

{body}

---
Write only the reply body. No subject line. Match my style from the examples above.
Reply:"""

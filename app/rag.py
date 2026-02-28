"""
MailMind — RAG Engine
Stores past emails as embeddings. Retrieves similar threads
at draft-time to match writing style and tone.
"""

import json
import logging
import os
import re
from collections import Counter
from typing import Optional

import chromadb
from chromadb.utils import embedding_functions

from .config import settings

logger = logging.getLogger("mailmind.rag")

STYLE_PROFILE_FILE = os.path.join(settings.DATA_DIR, "style_profile.json")


class RAGEngine:
    def __init__(self):
        os.makedirs(settings.VECTOR_DB_DIR, exist_ok=True)
        os.makedirs(settings.DATA_DIR, exist_ok=True)

        self.client = chromadb.PersistentClient(path=settings.VECTOR_DB_DIR)

        # Use sentence-transformers for local embeddings (no API calls)
        self.ef = embedding_functions.SentenceTransformerEmbeddingFunction(
            model_name="all-MiniLM-L6-v2"  # Small, fast, runs locally
        )

        self.collection = self.client.get_or_create_collection(
            name="email_history",
            embedding_function=self.ef,
            metadata={"hnsw:space": "cosine"}
        )

        logger.info(f"RAG engine ready. {self.count()} emails indexed.")

    def count(self) -> int:
        return self.collection.count()

    # ─── INGEST ──────────────────────────────────────────────────────────────

    def ingest(self, emails: list[dict]) -> int:
        """
        Index a list of sent emails into the vector store.
        Each email should have: id, subject, body, from, to
        """
        documents = []
        metadatas = []
        ids = []

        for email in emails[:settings.MAX_EMAIL_HISTORY]:
            email_id = str(email.get("id", ""))
            subject = email.get("subject", "")
            body = email.get("body", "")[:800]
            received = email.get("received", "")[:400]

            if not body:
                continue

            # Document = what we embed (for similarity search)
            doc = f"Subject: {subject}\n\nBody: {body}"

            documents.append(doc)
            metadatas.append({
                "id": email_id,
                "subject": subject,
                "reply": body[:500],
                "received": received[:400],
                "from": str(email.get("from", "")),
                "to": str(email.get("to", "")),
            })
            ids.append(f"email_{email_id}" if email_id else f"email_{len(ids)}")

        if documents:
            # Upsert in batches of 100
            batch_size = 100
            for i in range(0, len(documents), batch_size):
                self.collection.upsert(
                    documents=documents[i:i+batch_size],
                    metadatas=metadatas[i:i+batch_size],
                    ids=ids[i:i+batch_size]
                )

        logger.info(f"Ingested {len(documents)} emails into RAG store")
        return len(documents)

    # ─── RETRIEVE ────────────────────────────────────────────────────────────

    def retrieve(self, query: str, n_results: int = 5) -> list[dict]:
        """
        Find past emails most similar to the current incoming email.
        Returns metadata with 'received' and 'reply' fields for context.
        """
        if self.count() == 0:
            return []

        try:
            results = self.collection.query(
                query_texts=[query],
                n_results=min(n_results, self.count()),
                include=["metadatas", "distances"]
            )

            contexts = []
            for meta, dist in zip(
                results["metadatas"][0],
                results["distances"][0]
            ):
                # Only include if similarity is decent (distance < 0.8)
                if dist < 0.8:
                    contexts.append({
                        "received": meta.get("received", ""),
                        "reply": meta.get("reply", ""),
                        "subject": meta.get("subject", ""),
                        "similarity": round(1 - dist, 3)
                    })

            return contexts

        except Exception as e:
            logger.error(f"RAG retrieval error: {e}")
            return []

    # ─── STYLE PROFILE ───────────────────────────────────────────────────────

    def build_style_profile(self, emails: list[dict]) -> dict:
        """
        Analyze sent emails to extract writing style patterns.
        Saves to disk so it persists between sessions.
        """
        bodies = [e.get("body", "") for e in emails if e.get("body")]
        if not bodies:
            return {}

        # Detect sign-off style
        sign_offs = []
        sign_off_patterns = [
            r"(Best regards|Kind regards|Thanks|Thank you|Cheers|Regards|Sincerely|Talk soon|Speak soon|All the best)[,\s]",
        ]
        for body in bodies:
            for pattern in sign_off_patterns:
                match = re.search(pattern, body, re.IGNORECASE)
                if match:
                    sign_offs.append(match.group(1))

        most_common_signoff = Counter(sign_offs).most_common(1)
        sign_off = most_common_signoff[0][0] if most_common_signoff else "Best regards"

        # Detect formality
        formal_words = sum(1 for b in bodies for w in ["please", "kindly", "Dear", "hereby", "enclosed"] if w.lower() in b.lower())
        casual_words = sum(1 for b in bodies for w in ["hey", "hi", "thanks!", "yeah", "sure", "cool"] if w.lower() in b.lower())
        tone = "formal" if formal_words > casual_words else "conversational"

        # Average length
        avg_words = sum(len(b.split()) for b in bodies) / len(bodies)
        if avg_words < 50:
            avg_length = "very short (1-2 sentences)"
        elif avg_words < 120:
            avg_length = "concise (3-5 sentences)"
        elif avg_words < 300:
            avg_length = "moderate (1-2 paragraphs)"
        else:
            avg_length = "detailed (multiple paragraphs)"

        # Common phrases (2-3 word n-grams that appear multiple times)
        all_words = " ".join(bodies[:50]).lower()
        words = re.findall(r'\b[a-z]{3,}\b', all_words)
        bigrams = [f"{words[i]} {words[i+1]}" for i in range(len(words)-1)]
        common = [p for p, c in Counter(bigrams).most_common(20) if c >= 2][:8]

        profile = {
            "tone": tone,
            "avg_length": avg_length,
            "sign_off": sign_off,
            "common_phrases": common,
            "email_count": len(emails)
        }

        # Persist
        with open(STYLE_PROFILE_FILE, "w") as f:
            json.dump(profile, f, indent=2)

        logger.info(f"Style profile built: {profile}")
        return profile

    def get_style_profile(self) -> dict:
        """Load style profile from disk."""
        if os.path.exists(STYLE_PROFILE_FILE):
            with open(STYLE_PROFILE_FILE) as f:
                return json.load(f)
        return {}

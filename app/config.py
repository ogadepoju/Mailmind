"""MailMind Hosted Backend — Config"""
from pydantic_settings import BaseSettings

class Settings(BaseSettings):
    # Anthropic
    ANTHROPIC_API_KEY: str = ""
    ANTHROPIC_MODEL: str = "claude-haiku-4-5-20251001"

    # RAG
    DATA_DIR: str = "./data"
    VECTOR_DB_DIR: str = "./data/vectordb"
    MAX_RAG_RESULTS: int = 5
    MAX_EMAIL_HISTORY: int = 500

    class Config:
        env_file = ".env"
        env_file_encoding = "utf-8"

settings = Settings()

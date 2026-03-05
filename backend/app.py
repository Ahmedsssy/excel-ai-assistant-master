"""
Excel AI Assistant – Python Backend Server
==========================================
FastAPI server that:
  1. Proxies chat requests to Ollama / LM Studio (OpenAI-compatible API)
  2. Performs server-side statistical analysis (NumPy / SciPy / pandas)
  3. Detects anomalies (IQR, Z-score, Isolation Forest)
  4. Detects trends (linear regression, seasonal decomposition)

All data stays on the local machine – no external calls are made.

Run with:
    uvicorn app:app --host 0.0.0.0 --port 8000 --reload
"""

from __future__ import annotations

import json
import logging
from typing import Any

import httpx
import numpy as np
import pandas as pd
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
from pydantic import BaseModel, Field

from analyzer import (
    analyze_dataframe,
    detect_anomalies,
    detect_trends,
)

# ──────────────────────────────────────────────
#  App setup
# ──────────────────────────────────────────────

logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
logger = logging.getLogger(__name__)

app = FastAPI(
    title="Excel AI Assistant Backend",
    description="Local AI proxy + advanced data analysis for the Excel add-in.",
    version="1.0.0",
)

# Allow the Office add-in (running on localhost:3000) to call this server
app.add_middleware(
    CORSMiddleware,
    allow_origins=[
        "http://localhost:3000",
        "https://localhost:3000",
        "http://127.0.0.1:3000",
        "https://127.0.0.1:3000",
    ],
    allow_methods=["GET", "POST"],
    allow_headers=["Content-Type"],
)

# ──────────────────────────────────────────────
#  Pydantic models
# ──────────────────────────────────────────────


class Message(BaseModel):
    role: str
    content: str


class ChatRequest(BaseModel):
    messages: list[Message]
    model: str = "local-model"
    max_tokens: int = Field(default=2048, ge=1, le=32768)
    temperature: float = Field(default=0.7, ge=0.0, le=2.0)
    stream: bool = True
    ai_backend_url: str = "http://localhost:1234/v1"


class DataRequest(BaseModel):
    """Raw 2-D data from the Excel selection."""
    headers: list[str | None] | None = None
    rows: list[list[Any]]


class AnomalyRequest(DataRequest):
    z_threshold: float = Field(default=3.0, ge=1.0, le=6.0)
    use_isolation_forest: bool = False


class TrendRequest(DataRequest):
    pass


# ──────────────────────────────────────────────
#  Health / status
# ──────────────────────────────────────────────


@app.get("/health")
async def health():
    return {"status": "ok", "service": "excel-ai-assistant-backend"}


# ──────────────────────────────────────────────
#  AI proxy  (streaming & non-streaming)
# ──────────────────────────────────────────────


@app.post("/api/chat")
async def proxy_chat(req: ChatRequest):
    """
    Forward a chat request to the local AI backend (LM Studio / Ollama).
    Supports streaming (Server-Sent Events).
    """
    payload = {
        "model": req.model,
        "messages": [m.model_dump() for m in req.messages],
        "max_tokens": req.max_tokens,
        "temperature": req.temperature,
        "stream": req.stream,
    }

    target_url = f"{req.ai_backend_url.rstrip('/')}/chat/completions"

    if req.stream:
        async def stream_generator():
            async with httpx.AsyncClient(timeout=120) as client:
                try:
                    async with client.stream("POST", target_url, json=payload) as resp:
                        if resp.status_code != 200:
                            err = await resp.aread()
                            yield f"data: {json.dumps({'error': err.decode()})}\n\n"
                            return
                        async for line in resp.aiter_lines():
                            if line:
                                yield f"{line}\n\n"
                except httpx.ConnectError:
                    yield f"data: {json.dumps({'error': 'Cannot connect to AI backend'})}\n\n"

        return StreamingResponse(stream_generator(), media_type="text/event-stream")
    else:
        async with httpx.AsyncClient(timeout=120) as client:
            try:
                resp = await client.post(target_url, json=payload)
                resp.raise_for_status()
                return resp.json()
            except httpx.ConnectError:
                raise HTTPException(502, "Cannot connect to AI backend")
            except httpx.HTTPStatusError as e:
                raise HTTPException(e.response.status_code, str(e))


# ──────────────────────────────────────────────
#  Data analysis
# ──────────────────────────────────────────────


@app.post("/api/analyze")
async def analyze_data(req: DataRequest):
    """
    Perform descriptive statistics on the provided data.
    Returns per-column statistics and a text report.
    """
    try:
        df = _to_dataframe(req)
        result = analyze_dataframe(df)
        return {"ok": True, "analysis": result}
    except Exception as exc:
        logger.exception("analyze_data failed")
        raise HTTPException(500, str(exc))


@app.post("/api/anomalies")
async def find_anomalies(req: AnomalyRequest):
    """
    Detect outliers / anomalies using IQR, Z-score, and (optionally)
    Isolation Forest.
    """
    try:
        df = _to_dataframe(req)
        anomalies = detect_anomalies(
            df,
            z_threshold=req.z_threshold,
            use_isolation_forest=req.use_isolation_forest,
        )
        return {"ok": True, "anomalies": anomalies}
    except Exception as exc:
        logger.exception("find_anomalies failed")
        raise HTTPException(500, str(exc))


@app.post("/api/trends")
async def find_trends(req: TrendRequest):
    """
    Detect linear trends in numeric columns using linear regression.
    """
    try:
        df = _to_dataframe(req)
        trends = detect_trends(df)
        return {"ok": True, "trends": trends}
    except Exception as exc:
        logger.exception("find_trends failed")
        raise HTTPException(500, str(exc))


# ──────────────────────────────────────────────
#  Internal helpers
# ──────────────────────────────────────────────


def _to_dataframe(req: DataRequest) -> pd.DataFrame:
    """Convert the incoming request data to a pandas DataFrame."""
    if not req.rows:
        raise ValueError("No rows provided")

    df = pd.DataFrame(req.rows, columns=req.headers if req.headers else None)
    # Attempt numeric coercion on object columns
    for col in df.columns:
        try:
            df[col] = pd.to_numeric(df[col])
        except (ValueError, TypeError):
            pass
    return df

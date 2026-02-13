from __future__ import annotations

import os
import time
import wave
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, List, Optional

import pandas as pd
import requests
from tqdm import tqdm

def fetch_shared_voices(language: str = "he", page_size: int = 100, max_pages: int = 50) -> List[Dict[str, Any]]:
    """
    Fetch voices from ElevenLabs Voice Library (shared voices).
    Requires plan/API key access to Voice Library via API.
    """
    url = f"{ELEVENLABS_API_BASE}/v1/shared-voices"
    params: Dict[str, Any] = {"page_size": page_size, "language": language}

    voices: List[Dict[str, Any]] = []
    next_page_token: Optional[str] = None
    pages = 0

    while pages < max_pages:
        if next_page_token:
            params["next_page_token"] = next_page_token

        r = requests.get(url, headers=_headers(), params=params, timeout=60)

        # If your plan doesn't allow Voice Library API access, this is the common failure.
        if r.status_code == 403:
            raise RuntimeError(
                "403 Forbidden calling /v1/shared-voices. "
                "Your plan/API key likely doesn't have Voice Library API access."
            )

        r.raise_for_status()
        data = r.json()

        batch = data.get("voices") or data.get("shared_voices") or []
        voices.extend(batch)

        if not data.get("has_more"):
            break

        next_page_token = data.get("next_page_token")
        if not next_page_token:
            break

        pages += 1

    return voices


def export_shared_voices_to_excel(shared_voices: List[Dict[str, Any]], out_xlsx: Path) -> None:
    """
    Normalizes shared voices into a flat table and exports to Excel.
    """
    if out_xlsx.suffix.lower() not in [".xlsx", ".xlsm"]:
        raise ValueError("Excel output must be .xlsx or .xlsm")

    rows = []
    for v in shared_voices:
        labels = v.get("labels") or {}
        rows.append(
            {
                "voice_id": v.get("voice_id") or v.get("id"),
                "name": v.get("name"),
                "gender": labels.get("gender"),
                "accent": labels.get("accent"),
                "age": labels.get("age"),
                "category": v.get("category"),
                "description": v.get("description"),
                "preview_url": v.get("preview_url"),
                "language": v.get("language") or labels.get("language"),
            }
        )

    df = pd.DataFrame(rows)
    if not df.empty:
        df = df.sort_values(["name"], na_position="last")

    out_xlsx.parent.mkdir(parents=True, exist_ok=True)
    df.to_excel(out_xlsx, index=False, sheet_name="shared_voices")


# -----------------------------
# CONFIG (LOCKED TO YOUR SPECS)
# -----------------------------
ELEVENLABS_API_BASE = "https://api.elevenlabs.io"

# Request raw PCM S16LE @ 48kHz from ElevenLabs, then wrap into WAV
LOCKED_OUTPUT_FORMAT = "pcm_48000"

WAV_SAMPLE_RATE = 48000
WAV_SAMPLE_WIDTH_BYTES = 2   # 16-bit = 2 bytes (S16LE)
WAV_CHANNELS = 1             # typically mono for ElevenLabs PCM; adjust if your outputs are stereo

DEFAULT_MODEL_ID = "eleven_multilingual_v2"


# -----------------------------
# API KEY / HEADERS
# -----------------------------
def _get_api_key() -> str:
    key = os.getenv("ELEVENLABS_API_KEY") or os.getenv("XI_API_KEY")
    if not key:
        raise RuntimeError(
            "Missing API key. Set ELEVENLABS_API_KEY (recommended) or XI_API_KEY."
        )
    return key


def _headers() -> Dict[str, str]:
    return {"xi-api-key": _get_api_key()}


# -----------------------------
# 1) LIST VOICES -> EXCEL ONLY
# -----------------------------
def fetch_all_voices(page_size: int = 100) -> List[Dict[str, Any]]:
    """
    Fetch voices using v2 voices endpoint with pagination.
    Extract gender from labels.gender if present (or sharing.labels.gender).
    """
    url = f"{ELEVENLABS_API_BASE}/v2/voices"
    params: Dict[str, Any] = {"page_size": page_size, "include_total_count": True}

    voices: List[Dict[str, Any]] = []
    next_page_token: Optional[str] = None

    while True:
        if next_page_token:
            params["next_page_token"] = next_page_token

        r = requests.get(url, headers=_headers(), params=params, timeout=60)
        r.raise_for_status()
        data = r.json()

        for v in data.get("voices", []):
            labels = v.get("labels") or {}
            sharing_labels = (v.get("sharing") or {}).get("labels") or {}

            def pick_label(k: str) -> Optional[str]:
                return labels.get(k) or sharing_labels.get(k)

            voices.append(
                {
                    "voice_id": v.get("voice_id"),
                    "name": v.get("name"),
                    "category": v.get("category"),
                    "gender": pick_label("gender"),
                    "accent": pick_label("accent"),
                    "age": pick_label("age"),
                    "description": v.get("description"),
                    "preview_url": v.get("preview_url"),
                    "is_legacy": v.get("is_legacy"),
                    "is_owner": v.get("is_owner"),
                    "created_at_unix": v.get("created_at_unix"),
                }
            )

        if not data.get("has_more"):
            break

        next_page_token = data.get("next_page_token")
        if not next_page_token:
            break  # defensive

    return voices


def export_voices_to_excel(voices: List[Dict[str, Any]], out_xlsx: Path) -> None:
    """
    Excel only. Writes a single sheet named 'voices'.
    """
    if out_xlsx.suffix.lower() not in [".xlsx", ".xlsm"]:
        raise ValueError("Voices export must be an Excel file: .xlsx or .xlsm")

    out_xlsx.parent.mkdir(parents=True, exist_ok=True)
    df = pd.DataFrame(voices).sort_values(["category", "name"], na_position="last")
    df.to_excel(out_xlsx, index=False, sheet_name="voices")


# -----------------------------
# 2) EXCEL JOBS -> WAV ONLY
# -----------------------------
@dataclass
class TTSJob:
    filename: str          # forced to .wav
    voice_id: str
    text: str
    model_id: str = DEFAULT_MODEL_ID
    language_code: Optional[str] = None
    stability: Optional[float] = None
    similarity_boost: Optional[float] = None
    style: Optional[float] = None
    use_speaker_boost: Optional[bool] = None
    seed: Optional[int] = None


def _is_blank(x: Any) -> bool:
    if x is None:
        return True
    if isinstance(x, float) and pd.isna(x):
        return True
    return str(x).strip() == ""


def _coerce_bool(x: Any) -> Optional[bool]:
    if _is_blank(x):
        return None
    if isinstance(x, bool):
        return x
    s = str(x).strip().lower()
    if s in ("true", "1", "yes", "y"):
        return True
    if s in ("false", "0", "no", "n"):
        return False
    return None


def _coerce_float(x: Any) -> Optional[float]:
    if _is_blank(x):
        return None
    try:
        return float(x)
    except Exception:
        return None


def _coerce_int(x: Any) -> Optional[int]:
    if _is_blank(x):
        return None
    try:
        return int(float(x))
    except Exception:
        return None


def load_jobs_from_excel(xlsx_path: Path, sheet_name: str = "tts_jobs") -> List[TTSJob]:
    """
    Excel only to avoid CSV delimiter issues.
    Required columns: filename, voice_id, text
    Optional: model_id, language_code, stability, similarity_boost, style, use_speaker_boost, seed
    """
    if not xlsx_path.exists():
        raise FileNotFoundError(xlsx_path)
    if xlsx_path.suffix.lower() not in [".xlsx", ".xlsm"]:
        raise ValueError("Jobs input must be an Excel file: .xlsx or .xlsm")

    df = pd.read_excel(xlsx_path, sheet_name=sheet_name, dtype=object).fillna("")

    required = ["filename", "voice_id", "text"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"Jobs sheet missing required columns: {missing}")

    jobs: List[TTSJob] = []
    for _, row in df.iterrows():
        if _is_blank(row["filename"]) or _is_blank(row["voice_id"]) or _is_blank(row["text"]):
            continue

        filename = str(row["filename"]).strip()
        if not filename.lower().endswith(".wav"):
            filename += ".wav"

        jobs.append(
            TTSJob(
                filename=filename,
                voice_id=str(row["voice_id"]).strip(),
                text=str(row["text"]),
                model_id=(str(row.get("model_id", "")).strip() or DEFAULT_MODEL_ID),
                language_code=(str(row.get("language_code", "")).strip() or None),
                stability=_coerce_float(row.get("stability")),
                similarity_boost=_coerce_float(row.get("similarity_boost")),
                style=_coerce_float(row.get("style")),
                use_speaker_boost=_coerce_bool(row.get("use_speaker_boost")),
                seed=_coerce_int(row.get("seed")),
            )
        )

    return jobs


def _wrap_pcm_to_wav(pcm_bytes: bytes, wav_path: Path) -> None:
    """
    Wrap raw PCM S16LE into a WAV container (48kHz, 16-bit).
    """
    wav_path.parent.mkdir(parents=True, exist_ok=True)
    with wave.open(str(wav_path), "wb") as wf:
        wf.setnchannels(WAV_CHANNELS)
        wf.setsampwidth(WAV_SAMPLE_WIDTH_BYTES)
        wf.setframerate(WAV_SAMPLE_RATE)
        wf.writeframes(pcm_bytes)


def synthesize_one_to_wav(job: TTSJob, out_dir: Path, retries: int = 3) -> Path:
    """
    Calls ElevenLabs with output_format=pcm_48000, then wraps into WAV.
    """
    url = f"{ELEVENLABS_API_BASE}/v1/text-to-speech/{job.voice_id}"
    params = {"output_format": LOCKED_OUTPUT_FORMAT}  # locked

    payload: Dict[str, Any] = {"text": job.text, "model_id": job.model_id}

    if job.language_code:
        payload["language_code"] = job.language_code

    voice_settings: Dict[str, Any] = {}
    if job.stability is not None:
        voice_settings["stability"] = job.stability
    if job.similarity_boost is not None:
        voice_settings["similarity_boost"] = job.similarity_boost
    if job.style is not None:
        voice_settings["style"] = job.style
    if job.use_speaker_boost is not None:
        voice_settings["use_speaker_boost"] = job.use_speaker_boost
    if voice_settings:
        payload["voice_settings"] = voice_settings

    if job.seed is not None:
        payload["seed"] = job.seed

    out_dir.mkdir(parents=True, exist_ok=True)
    wav_path = out_dir / job.filename

    backoff = 1.0
    for attempt in range(1, retries + 1):
        try:
            r = requests.post(
                url,
                headers={**_headers(), "Content-Type": "application/json"},
                params=params,
                json=payload,
                timeout=180,
            )

            if r.status_code == 429 or 500 <= r.status_code < 600:
                raise RuntimeError(f"HTTP {r.status_code}: {r.text[:300]}")

            r.raise_for_status()

            pcm_bytes = r.content
            _wrap_pcm_to_wav(pcm_bytes, wav_path)
            return wav_path

        except Exception as e:
            if attempt < retries:
                time.sleep(backoff)
                backoff *= 2
            else:
                raise RuntimeError(f"Failed job {job.filename}: {e}") from e

    raise RuntimeError("Unexpected failure")


def run_batch_from_excel(
    xlsx_path: Path,
    out_dir: Path,
    sheet_name: str = "tts_jobs",
    zip_outputs: bool = True,
) -> Path:
    jobs = load_jobs_from_excel(xlsx_path, sheet_name=sheet_name)
    if not jobs:
        raise RuntimeError("No valid jobs found (check filename/voice_id/text columns).")

    written: List[Path] = []
    for job in tqdm(jobs, desc="Generating WAV (48kHz 16-bit PCM)"):
        p = synthesize_one_to_wav(job, out_dir=out_dir)
        written.append(p)

    if zip_outputs:
        zip_path = out_dir.with_suffix(".zip")
        with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as z:
            for p in written:
                z.write(p, arcname=p.name)
        return zip_path

    return out_dir


# -----------------------------
# CLI
# -----------------------------
if __name__ == "__main__":
    import argparse

    ap = argparse.ArgumentParser(
        description="ElevenLabs: export voices to Excel + batch TTS to WAV (48kHz, 16-bit PCM S16LE)."
    )
    ap.add_argument("--export-voices-xlsx", type=str, default="")
    ap.add_argument("--jobs-xlsx", type=str, default="")
    ap.add_argument("--jobs-sheet", type=str, default="tts_jobs")
    ap.add_argument("--out", type=str, default="out_audio")
    ap.add_argument("--zip", action="store_true")
    ap.add_argument("--export-shared-voices-xlsx", type=str, default="")
    ap.add_argument("--shared-language", type=str, default="he")

    args = ap.parse_args()

    if args.export_voices_xlsx:
        voices = fetch_all_voices()
        export_voices_to_excel(voices, Path(args.export_voices_xlsx))
        print(f"Wrote voices Excel to: {args.export_voices_xlsx}")

    if args.export_shared_voices_xlsx:
        shared = fetch_shared_voices(language=args.shared_language)
        export_shared_voices_to_excel(shared, Path(args.export_shared_voices_xlsx))
        print(f"Wrote shared voices Excel to: {args.export_shared_voices_xlsx} (language={args.shared_language})")


    if args.jobs_xlsx:
        result = run_batch_from_excel(
            xlsx_path=Path(args.jobs_xlsx),
            out_dir=Path(args.out),
            sheet_name=args.jobs_sheet,
            zip_outputs=args.zip,
        )
        print(f"Batch complete: {result}")

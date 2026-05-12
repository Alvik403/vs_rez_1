from __future__ import annotations

import asyncio
import os
from collections.abc import Iterator
from concurrent.futures import ProcessPoolExecutor
from datetime import datetime
from pathlib import Path

from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.responses import HTMLResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles

from app.excel_processor import ParsedFile, build_consolidated_workbook, parse_reserve_workbook

BASE_DIR = Path(__file__).resolve().parent

# Не кэшировать ответ с данными пользователя (ни прокси, ни браузер)
_NO_STORE = {
    "Cache-Control": "no-store, no-cache, must-revalidate, private",
    "Pragma": "no-cache",
    "Expires": "0",
}

_STREAM_CHUNK = 256 * 1024


def _parse_worker_count() -> int:
    raw = os.getenv("RESERVE_PARSE_WORKERS")
    if raw:
        try:
            return max(1, int(raw))
        except ValueError:
            pass
    return max(1, min(2, os.cpu_count() or 1))


_PARSE_WORKERS = _parse_worker_count()
_PARSE_POOL: ProcessPoolExecutor | None = (
    ProcessPoolExecutor(max_workers=_PARSE_WORKERS) if _PARSE_WORKERS > 1 else None
)


def _parse_uploaded_excel(item: tuple[str, bytes]) -> ParsedFile:
    filename, content = item
    return parse_reserve_workbook(content, filename)


async def _parse_excel_files(items: list[tuple[str, bytes]]) -> tuple[list[ParsedFile], list[str]]:
    if not items:
        return [], []

    if _PARSE_POOL is None or len(items) == 1:
        parsed: list[ParsedFile] = []
        errors: list[str] = []
        for filename, content in items:
            try:
                parsed.append(parse_reserve_workbook(content, filename))
            except Exception as exc:
                errors.append(str(exc))
        return parsed, errors

    loop = asyncio.get_running_loop()
    tasks = [loop.run_in_executor(_PARSE_POOL, _parse_uploaded_excel, item) for item in items]
    results = await asyncio.gather(*tasks, return_exceptions=True)

    parsed = []
    errors = []
    for (filename, _content), result in zip(items, results, strict=True):
        if isinstance(result, Exception):
            errors.append(f"{filename}: {result}")
        else:
            parsed.append(result)

    return parsed, errors


def _iter_excel_bytes(payload: bytes) -> Iterator[bytes]:
    """Стрим по чанкам; memoryview в конце снимает ссылку на буфер представления."""
    if not payload:
        return
    mv = memoryview(payload)
    try:
        n = len(mv)
        i = 0
        while i < n:
            j = min(i + _STREAM_CHUNK, n)
            yield mv[i:j].tobytes()
            i = j
    finally:
        mv.release()


app = FastAPI(title="Консолидация резервов")
app.mount("/static", StaticFiles(directory=BASE_DIR / "static"), name="static")


@app.on_event("shutdown")
def shutdown_parse_pool() -> None:
    if _PARSE_POOL is not None:
        _PARSE_POOL.shutdown(cancel_futures=True)


@app.get("/", response_class=HTMLResponse)
async def index() -> HTMLResponse:
    body = (BASE_DIR / "templates" / "index.html").read_text(encoding="utf-8")
    return HTMLResponse(content=body, headers=dict(_NO_STORE))


@app.post("/process")
async def process_files(files: list[UploadFile] = File(...)) -> StreamingResponse:
    if not files:
        raise HTTPException(status_code=400, detail="Загрузите хотя бы один файл .xlsx или .xlsm.")

    upload_items: list[tuple[str, bytes]] = []
    errors: list[str] = []

    for uploaded in files:
        if not uploaded.filename.lower().endswith((".xlsx", ".xlsm")):
            errors.append(f"{uploaded.filename}: неподдерживаемый формат")
            continue
        content = await uploaded.read()
        upload_items.append((uploaded.filename, content))

    parsed_files, parse_errors = await _parse_excel_files(upload_items)
    errors.extend(parse_errors)
    upload_items.clear()

    if not parsed_files:
        raise HTTPException(status_code=400, detail="Не удалось прочитать файлы. " + "; ".join(errors))

    try:
        output, projects, openings = build_consolidated_workbook(parsed_files)
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"Ошибка сборки файла: {exc}") from exc
    finally:
        parsed_files.clear()

    project_n, opening_n = len(projects), len(openings)
    del projects, openings

    filename = f"consolidated_reserves_{datetime.now():%Y-%m-%d_%H-%M}.xlsx"
    media_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    headers = {
        **_NO_STORE,
        "Content-Disposition": f'attachment; filename="{filename}"',
        "X-Projects-Count": str(project_n),
        "X-Openings-Count": str(opening_n),
        "X-Output-Ext": "xlsx",
    }
    if errors:
        headers["X-Import-Warnings"] = " | ".join(errors)

    return StreamingResponse(
        _iter_excel_bytes(output),
        media_type=media_type,
        headers=headers,
    )


@app.get("/health")
async def health() -> dict[str, str]:
    return {"status": "ok"}

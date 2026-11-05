"""
main.py — FastAPI приложение для маскировки персональных данных в Excel.
Сборка: pyinstaller --onefile --noconsole --add-data "index.html;." main.py
"""

# ============================================================
# МОДУЛЬ: ИМПОРТЫ И КОНСТАНТЫ
# ============================================================
import os
import sys
import json
import hashlib
import logging
import shutil
import threading
import time
from pathlib import Path
from datetime import datetime
from typing import Optional, Dict, Any, List, Tuple

from fastapi import FastAPI, UploadFile, File, Form, HTTPException, Depends, Request
from fastapi.responses import HTMLResponse, FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from sqlalchemy import create_engine, Column, String, Integer, DateTime, Text, ForeignKey
from sqlalchemy.orm import declarative_base, sessionmaker, Session
import openpyxl
from openpyxl.utils import get_column_letter
from faker import Faker

if getattr(sys, 'frozen', False):
    BASE_DIR = Path(sys._MEIPASS)
    DATA_DIR = Path(sys.executable).parent / "data"
else:
    BASE_DIR = Path(__file__).parent
    DATA_DIR = BASE_DIR / "data"

DATA_DIR.mkdir(exist_ok=True)

USERS_FILE = DATA_DIR / "users.json"
DB_FILE = DATA_DIR / "progress.db"
UPLOAD_DIR = DATA_DIR / "uploads"
MASKED_DIR = DATA_DIR / "masked"
UPLOAD_DIR.mkdir(exist_ok=True)
MASKED_DIR.mkdir(exist_ok=True)

LOG_FILE = DATA_DIR / "activity.log"
logging.basicConfig(
    filename=str(LOG_FILE),
    level=logging.INFO,
    format="%(asctime)s | %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
    encoding='utf-8'
)
logger = logging.getLogger("activity")

fake = Faker('ru_RU')

# ============================================================
# МОДУЛЬ: УПРАВЛЕНИЕ ПОЛЬЗОВАТЕЛЯМИ (JSON, ДИНАМИЧЕСКОЕ ЧТЕНИЕ)
# ============================================================
def ensure_users_file():
    """Создаёт users.json с admin/admin, если файла нет."""
    if not USERS_FILE.exists():
        default = {
            "users": [
                {"login": "admin", "password": "admin"}
            ]
        }
        with open(USERS_FILE, 'w', encoding='utf-8') as f:
            json.dump(default, f, ensure_ascii=False, indent=2)
        logger.info("Создан файл users.json с учёткой admin/admin")

def load_users() -> List[Dict[str, str]]:
    """Читает users.json при каждом вызове (динамически)."""
    ensure_users_file()
    try:
        with open(USERS_FILE, 'r', encoding='utf-8') as f:
            data = json.load(f)
        return data.get("users", [])
    except Exception as e:
        logger.error(f"Ошибка чтения users.json: {e}")
        return []

def authenticate(login: str, password: str) -> bool:
    """Проверяет пару логин/пароль по актуальному файлу."""
    users = load_users()
    for u in users:
        if u.get("login") == login and u.get("password") == password:
            return True
    return False

def user_exists(login: str) -> bool:
    users = load_users()
    return any(u.get("login") == login for u in users)

# ============================================================
# МОДУЛЬ: БАЗА ДАННЫХ SQLITE (ПРОГРЕСС МАСКИРОВКИ)
# ============================================================
Base = declarative_base()

class MaskRecord(Base):
    """Запись соответствия: оригинальное значение -> замаскированное.
    Ключ: (user_login, original_hash, column_name, original_value)"""
    __tablename__ = "mask_records"
    id = Column(Integer, primary_key=True, autoincrement=True)
    user_login = Column(String, index=True, nullable=False)
    original_hash = Column(String, index=True, nullable=False)
    column_name = Column(String, nullable=False)
    original_value = Column(Text, nullable=False)
    masked_value = Column(Text, nullable=False)
    created_at = Column(DateTime, default=datetime.utcnow)

class FileProgress(Base):
    """Отслеживание обработанных файлов (для логирования и статистики)."""
    __tablename__ = "file_progress"
    id = Column(Integer, primary_key=True, autoincrement=True)
    user_login = Column(String, index=True, nullable=False)
    original_hash = Column(String, index=True, nullable=False)
    original_filename = Column(String)
    processed_at = Column(DateTime, default=datetime.utcnow)
    status = Column(String, default="completed")

engine = create_engine(
    f"sqlite:///{DB_FILE}",
    connect_args={"check_same_thread": False}
)
Base.metadata.create_all(engine)
SessionLocal = sessionmaker(bind=engine, autoflush=False, autocommit=False)

def get_db():
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()

# ============================================================
# МОДУЛЬ: АНАЛИЗ EXCEL И МАСКИРОВКА
# ============================================================
def compute_file_hash(content: bytes) -> str:
    return hashlib.sha256(content).hexdigest()

def detect_column_type(values: List[Any]) -> str:
    """
    Эвристически определяет тип колонки по содержимому.
    Возвращает: 'fio', 'phone', 'passport', 'email', 'date', 'number', 'text'.
    """
    sample = [str(v).strip() for v in values if v is not None and str(v).strip() != ""]
    if not sample:
        return "text"
    
    fio_count = sum(1 for s in sample if len(s.split()) == 3 and all(p[0].isupper() for p in s.split() if p))
    if fio_count / len(sample) > 0.6:
        return "fio"
    
    import re
    phone_pattern = re.compile(r'^[\+]?[7|8][\s\-]?\(?\d{3}\)?[\s\-]?\d{3}[\s\-]?\d{2}[\s\-]?\d{2}$')
    if sum(1 for s in sample if phone_pattern.match(s.replace(" ", ""))) / len(sample) > 0.6:
        return "phone"
    
    passport_pattern = re.compile(r'^\d{4}[\s\-]?\d{6}$')
    if sum(1 for s in sample if passport_pattern.match(s.replace(" ", ""))) / len(sample) > 0.6:
        return "passport"
    
    if sum(1 for s in sample if "@" in s and "." in s) / len(sample) > 0.6:
        return "email"
    
    date_pattern = re.compile(r'\d{2}[\.\/\-]\d{2}[\.\/\-]\d{2,4}')
    if sum(1 for s in sample if date_pattern.search(s)) / len(sample) > 0.6:
        return "date"
    
    try:
        for s in sample:
            float(s.replace(",", "."))
        return "number"
    except ValueError:
        pass
    
    return "text"

def generate_fake_value(original: str, col_type: str) -> str:
    """Генерирует осмысленное фейковое значение в зависимости от типа.
    В конце подмешивает случайные буквы к результату."""
    original = str(original).strip()
    if not original:
        return original
    
    if col_type == "fio":
        value = fake.name()
    elif col_type == "phone":
        value = fake.phone_number()
    elif col_type == "passport":
        value = f"{fake.random_int(1000, 9999)} {fake.random_int(100000, 999999)}"
    elif col_type == "email":
        value = fake.email()
    elif col_type == "date":
        value = fake.date(pattern="%d.%m.%Y")
    elif col_type == "number":
        if "," in original:
            value = str(round(fake.random.uniform(1, 99999), 2)).replace(".", ",")
        else:
            value = str(fake.random_int(1, 99999))
    else:
        value = original
    
    value = _mix_random_chars(value)
    return value


CYRILLIC_LOWER = "абвгдеёжзийклмнопрстуфхцчшщъыьэюя"
CYRILLIC_UPPER = "АБВГДЕЁЖЗИЙКЛМНОПРСТУФХЦЧШЩЪЫЬЭЮЯ"
LATIN_LOWER = "abcdefghijklmnopqrstuvwxyz"
LATIN_UPPER = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
ALL_LETTERS = CYRILLIC_LOWER + CYRILLIC_UPPER + LATIN_LOWER + LATIN_UPPER


def _mix_random_chars(text: str, mix_rate: float = 0.15) -> str:
    """Подмешивает случайные буквы внутрь строки с заданной частотой."""
    import random
    if not text or len(text) <= 2:
        return text
    result = []
    for ch in text:
        result.append(ch)
        if random.random() < mix_rate:
            result.append(random.choice(ALL_LETTERS))
    return "".join(result)

def mask_excel(input_path: Path, output_path: Path, user_login: str, original_hash: str, db: Session) -> Dict[str, Any]:
    """
    Читает Excel, определяет типы колонок, генерирует замаскированные значения,
    сохраняет соответствия в SQLite и записывает новый Excel.
    Возвращает предпросмотр (первые 5 строк) и статистику.
    """
    wb = openpyxl.load_workbook(input_path)
    ws = wb.active
    
    headers = []
    rows = []
    for row_idx, row in enumerate(ws.iter_rows(values_only=True), start=1):
        if row_idx == 1:
            headers = [str(c) if c is not None else f"col_{i}" for i, c in enumerate(row)]
        else:
            rows.append(row)
    
    if not rows:
        raise ValueError("Файл пуст (нет данных после заголовка)")
    
    sample_rows = rows[:50]
    col_types = {}
    for col_idx, header in enumerate(headers):
        col_values = [r[col_idx] if col_idx < len(r) else None for r in sample_rows]
        col_types[header] = detect_column_type(col_values)
    
    new_wb = openpyxl.Workbook()
    new_ws = new_wb.active
    
    for col_idx, header in enumerate(headers, start=1):
        new_ws.cell(row=1, column=col_idx, value=header)
    
    preview_rows = []
    total_cells_masked = 0
    
    for row_idx, row in enumerate(rows, start=2):
        new_row = []
        for col_idx, header in enumerate(headers):
            original_val = row[col_idx] if col_idx < len(row) else None
            if original_val is None or str(original_val).strip() == "":
                new_row.append(original_val)
                continue
            
            existing = db.query(MaskRecord).filter_by(
                user_login=user_login,
                original_hash=original_hash,
                column_name=header,
                original_value=str(original_val)
            ).first()
            
            if existing:
                masked_val = existing.masked_value
            else:
                masked_val = generate_fake_value(str(original_val), col_types[header])
                record = MaskRecord(
                    user_login=user_login,
                    original_hash=original_hash,
                    column_name=header,
                    original_value=str(original_val),
                    masked_value=masked_val
                )
                db.add(record)
                total_cells_masked += 1
            
            new_row.append(masked_val)
            
            if row_idx <= 6:
                preview_rows.append({
                    "column": header,
                    "original": str(original_val)[:50],
                    "masked": str(masked_val)[:50],
                    "detected_type": col_types[header]
                })
        
        for col_idx, val in enumerate(new_row, start=1):
            new_ws.cell(row=row_idx, column=col_idx, value=val)
    
    db.commit()
    new_wb.save(output_path)
    
    progress = FileProgress(
        user_login=user_login,
        original_hash=original_hash,
        original_filename=input_path.name,
        status="masked"
    )
    db.add(progress)
    db.commit()
    
    logger.info(f"MASK | user={user_login} | file_hash={original_hash[:12]}... | cells_masked={total_cells_masked}")
    
    return {
        "preview": preview_rows,
        "col_types": col_types,
        "total_cells_masked": total_cells_masked,
        "rows_processed": len(rows)
    }

def unmask_excel(masked_path: Path, output_path: Path, user_login: str, original_hash: str, db: Session) -> Dict[str, Any]:
    """
    Обратная конвертация: читает замаскированный файл и заменяет значения
    на оригинальные из SQLite по ключу (user_login, original_hash, column_name, masked_value).
    """
    wb = openpyxl.load_workbook(masked_path)
    ws = wb.active
    
    headers = []
    rows = []
    for row_idx, row in enumerate(ws.iter_rows(values_only=True), start=1):
        if row_idx == 1:
            headers = [str(c) if c is not None else f"col_{i}" for i, c in enumerate(row)]
        else:
            rows.append(row)
    
    if not rows:
        raise ValueError("Файл пуст")
    
    records = db.query(MaskRecord).filter_by(
        user_login=user_login,
        original_hash=original_hash
    ).all()
    
    reverse_map: Dict[Tuple[str, str], str] = {}
    for rec in records:
        reverse_map[(rec.column_name, rec.masked_value)] = rec.original_value
    
    new_wb = openpyxl.Workbook()
    new_ws = new_wb.active
    
    for col_idx, header in enumerate(headers, start=1):
        new_ws.cell(row=1, column=col_idx, value=header)
    
    restored_count = 0
    for row_idx, row in enumerate(rows, start=2):
        for col_idx, header in enumerate(headers):
            masked_val = row[col_idx] if col_idx < len(row) else None
            if masked_val is None:
                new_ws.cell(row=row_idx, column=col_idx + 1, value=None)
                continue
            
            masked_str = str(masked_val).strip()
            original = reverse_map.get((header, masked_str))
            if original is not None:
                new_ws.cell(row=row_idx, column=col_idx + 1, value=original)
                restored_count += 1
            else:
                new_ws.cell(row=row_idx, column=col_idx + 1, value=masked_val)
    
    new_wb.save(output_path)
    
    progress = FileProgress(
        user_login=user_login,
        original_hash=original_hash,
        original_filename=masked_path.name,
        status="unmasked"
    )
    db.add(progress)
    db.commit()
    
    logger.info(f"UNMASK | user={user_login} | file_hash={original_hash[:12]}... | cells_restored={restored_count}")
    
    return {
        "restored_cells": restored_count,
        "rows_processed": len(rows)
    }

# ============================================================
# МОДУЛЬ: FASTAPI ПРИЛОЖЕНИЕ И МАРШРУТЫ
# ============================================================
app = FastAPI(title="PII Masking Tool")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

# --- Вспомогательные модели ---
class LoginRequest(BaseModel):
    login: str
    password: str

# --- Аутентификация (простая, без токенов, для демо) ---

def get_current_user(request: Request) -> str:
    """Извлекает логин и пароль из заголовков и проверяет по JSON."""
    login = request.headers.get("X-Login", "")
    password = request.headers.get("X-Password", "")
    if not login or not password:
        raise HTTPException(status_code=401, detail="Требуется авторизация")
    if not authenticate(login, password):
        raise HTTPException(status_code=401, detail="Неверный логин или пароль")
    return login

# --- Маршруты ---

@app.get("/", response_class=HTMLResponse)
async def root():
    """Отдаёт index.html."""
    html_path = BASE_DIR / "index.html"
    if not html_path.exists():
        html_path = Path(__file__).parent / "index.html"
    if html_path.exists():
        return HTMLResponse(content=html_path.read_text(encoding="utf-8"))
    return HTMLResponse("<h1>index.html не найден</h1>", status_code=404)

@app.post("/api/login")
async def login(req: LoginRequest):
    """Проверка логина/пароля. Возвращает успех или ошибку."""
    if authenticate(req.login, req.password):
        logger.info(f"LOGIN | user={req.login}")
        return {"ok": True, "login": req.login}
    logger.warning(f"LOGIN_FAIL | user={req.login}")
    raise HTTPException(status_code=401, detail="Неверный логин или пароль")

@app.post("/api/upload")
async def upload_file(
    file: UploadFile = File(...),
    user: str = Depends(get_current_user),
    db: Session = Depends(get_db)
):
    """
    Загружает Excel, вычисляет хэш, сохраняет оригинал и возвращает предпросмотр
    (без маскировки). Маскировка происходит при вызове /api/process.
    """
    if not file.filename.lower().endswith(('.xlsx', '.xls')):
        raise HTTPException(400, "Только Excel файлы (.xlsx, .xls)")
    
    content = await file.read()
    file_hash = compute_file_hash(content)
    
    original_path = UPLOAD_DIR / f"{user}_{file_hash[:16]}_{file.filename}"
    with open(original_path, "wb") as f:
        f.write(content)
    
    logger.info(f"UPLOAD | user={user} | file={file.filename} | hash={file_hash[:12]}...")
    
    try:
        wb = openpyxl.load_workbook(original_path)
        ws = wb.active
        preview_headers = []
        preview_rows = []
        for row_idx, row in enumerate(ws.iter_rows(values_only=True), start=1):
            if row_idx == 1:
                preview_headers = [str(c) if c is not None else "" for c in row]
            elif row_idx <= 6:
                preview_rows.append([str(c) if c is not None else "" for c in row])
            else:
                break
        wb.close()
    except Exception as e:
        raise HTTPException(400, f"Ошибка чтения Excel: {e}")
    
    return {
        "file_hash": file_hash,
        "filename": file.filename,
        "headers": preview_headers,
        "preview_rows": preview_rows
    }

@app.post("/api/process")
async def process_file(
    file_hash: str = Form(...),
    filename: str = Form(...),
    user: str = Depends(get_current_user),
    db: Session = Depends(get_db)
):
    """
    Запускает маскировку загруженного файла.
    Возвращает предпросмотр замаскированных данных и путь для скачивания.
    """
    original_path = UPLOAD_DIR / f"{user}_{file_hash[:16]}_{filename}"
    if not original_path.exists():
        raise HTTPException(404, "Файл не найден. Загрузите заново.")
    
    output_filename = f"masked_{file_hash[:12]}_{filename}"
    masked_path = MASKED_DIR / output_filename
    
    try:
        result = mask_excel(original_path, masked_path, user, file_hash, db)
    except Exception as e:
        logger.error(f"PROCESS_ERROR | user={user} | error={e}")
        raise HTTPException(500, f"Ошибка обработки: {e}")
    
    return {
        "preview": result["preview"],
        "col_types": result["col_types"],
        "total_cells_masked": result["total_cells_masked"],
        "rows_processed": result["rows_processed"],
        "masked_filename": output_filename,
        "download_url": f"/api/download/{output_filename}"
    }

@app.get("/api/download/{filename}")
async def download_file(filename: str, user: str = Depends(get_current_user)):
    """Скачивание замаскированного файла."""
    if not filename.startswith("masked_"):
        raise HTTPException(400, "Неверный файл")
    
    file_path = MASKED_DIR / filename
    if not file_path.exists():
        raise HTTPException(404, "Файл не найден")
    
    logger.info(f"DOWNLOAD | user={user} | file={filename}")
    return FileResponse(
        file_path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=filename
    )

@app.post("/api/unmask")
async def unmask_file(
    file: UploadFile = File(...),
    original_hash: str = Form(...),
    user: str = Depends(get_current_user),
    db: Session = Depends(get_db)
):
    """
    Обратная конвертация: загружаем замаскированный файл и original_hash,
    восстанавливаем оригинальные значения из SQLite.
    """
    if not file.filename.lower().endswith(('.xlsx', '.xls')):
        raise HTTPException(400, "Только Excel файлы")
    
    content = await file.read()
    masked_upload_path = UPLOAD_DIR / f"{user}_unmask_{file.filename}"
    with open(masked_upload_path, "wb") as f:
        f.write(content)
    
    output_filename = f"restored_{file.filename}"
    restored_path = MASKED_DIR / output_filename
    
    try:
        result = unmask_excel(masked_upload_path, restored_path, user, original_hash, db)
    except Exception as e:
        logger.error(f"UNMASK_ERROR | user={user} | error={e}")
        raise HTTPException(500, f"Ошибка восстановления: {e}")
    
    return {
        "restored_cells": result["restored_cells"],
        "rows_processed": result["rows_processed"],
        "download_url": f"/api/download/restored/{output_filename}"
    }

@app.get("/api/download/restored/{filename}")
async def download_restored(filename: str, user: str = Depends(get_current_user)):
    """Скачивание восстановленного файла."""
    file_path = MASKED_DIR / filename
    if not file_path.exists():
        raise HTTPException(404, "Файл не найден")
    logger.info(f"DOWNLOAD_RESTORED | user={user} | file={filename}")
    return FileResponse(
        file_path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=filename
    )

@app.get("/api/logs")
async def get_logs(user: str = Depends(get_current_user)):
    """
    Возвращает последние строки лога (только для admin).
    В реальном приложении добавьте проверку роли.
    """
    if user != "admin":
        raise HTTPException(403, "Доступ только для администратора")
    
    lines = []
    if LOG_FILE.exists():
        with open(LOG_FILE, 'r', encoding='utf-8', errors='replace') as f:
            lines = f.readlines()[-200:]
    return {"logs": [l.strip() for l in lines]}

# ============================================================
# МОДУЛЬ: ЗАПУСК (ДЛЯ РАЗРАБОТКИ / PYINSTALLER)
# ============================================================


class _NoConsoleStream:
    """Заглушка для sys.stdout / sys.stderr при сборке --noconsole.
    PyInstaller заменяет их на None, из-за чего uvicorn падает
    (ColorizeFormatter вызывает sys.stderr.isatty())."""

    def write(self, *args, **kwargs):
        return 0

    def flush(self, *args, **kwargs):
        return 0

    def isatty(self) -> bool:
        return False


def run_server():
    """Запуск uvicorn. Вызывается из main или из exe."""
    if sys.stdout is None:
        sys.stdout = _NoConsoleStream()
    if sys.stderr is None:
        sys.stderr = _NoConsoleStream()

    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=9321, log_level="info")

if __name__ == "__main__":
    ensure_users_file()
    run_server()
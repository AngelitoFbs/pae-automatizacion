"""Utilidades seguras para manejo de archivos temporales y validación"""
import tempfile
import os
import uuid
from pathlib import Path
from contextlib import contextmanager
from typing import BinaryIO, Generator
import logging

logger = logging.getLogger(__name__)

# Extensiones permitidas para archivos Excel
ALLOWED_EXCEL_EXTENSIONS = {'.xlsx', '.xlsm'}
ALLOWED_CSV_EXTENSIONS = {'.csv'}
MAX_FILE_SIZE_MB = 50
MAX_FILE_SIZE_BYTES = MAX_FILE_SIZE_MB * 1024 * 1024


@contextmanager
def secure_temp_file(suffix: str = '.xlsx', prefix: str = 'pae_') -> Generator[Path, None, None]:
    """
    Context manager seguro para archivos temporales.
    Genera nombres únicos impredecibles y garantiza limpieza.
    """
    # Generar nombre único con UUID para evitar race conditions
    unique_name = f"{prefix}{uuid.uuid4().hex[:12]}{suffix}"
    temp_dir = Path(tempfile.gettempdir())
    temp_path = temp_dir / unique_name
    
    try:
        yield temp_path
    finally:
        # Limpieza garantizada
        try:
            if temp_path.exists():
                temp_path.unlink()
        except Exception as e:
            logger.warning(f"No se pudo eliminar temp file {temp_path}: {e}")


@contextmanager
def secure_temp_files(count: int = 2, suffix: str = '.xlsx') -> Generator[list[Path], None, None]:
    """Context manager para múltiples archivos temporales seguros."""
    paths = []
    try:
        for i in range(count):
            unique_name = f"pae_{uuid.uuid4().hex[:12]}_{i}{suffix}"
            temp_dir = Path(tempfile.gettempdir())
            temp_path = temp_dir / unique_name
            paths.append(temp_path)
        yield paths
    finally:
        for p in paths:
            try:
                if p.exists():
                    p.unlink()
            except Exception as e:
                logger.warning(f"No se pudo eliminar temp file {p}: {e}")


def validate_excel_file(file_bytes: bytes, filename: str) -> tuple[bool, str]:
    """
    Valida que el archivo sea un Excel válido.
    Returns: (is_valid, error_message)
    """
    # Validar tamaño
    if len(file_bytes) > MAX_FILE_SIZE_BYTES:
        return False, f"Archivo demasiado grande (> {MAX_FILE_SIZE_MB}MB)"
    
    if len(file_bytes) == 0:
        return False, "Archivo vacío"
    
    # Validar extensión
    ext = Path(filename).suffix.lower()
    if ext not in ALLOWED_EXCEL_EXTENSIONS:
        return False, f"Extensión no permitida: {ext}. Use .xlsx o .xlsm"
    
    # Validar magic bytes (signature del archivo)
    # XLSX = PK\x03\x04 (ZIP)
    if not file_bytes.startswith(b'PK\x03\x04'):
        return False, "El archivo no es un Excel válido (formato incorrecto)"
    
    return True, ""


def validate_csv_file(file_bytes: bytes, filename: str) -> tuple[bool, str]:
    """Valida que el archivo sea un CSV válido."""
    if len(file_bytes) > MAX_FILE_SIZE_BYTES:
        return False, f"Archivo demasiado grande (> {MAX_FILE_SIZE_MB}MB)"
    
    if len(file_bytes) == 0:
        return False, "Archivo vacío"
    
    ext = Path(filename).suffix.lower()
    if ext not in ALLOWED_CSV_EXTENSIONS:
        return False, f"Extensión no permitida: {ext}. Use .csv"
    
    # Validar que sea texto UTF-8 decodificable
    try:
        file_bytes.decode('utf-8')
    except UnicodeDecodeError:
        try:
            file_bytes.decode('latin-1')
        except UnicodeDecodeError:
            return False, "El archivo no es un CSV válido (codificación no soportada)"
    
    return True, ""


def safe_save_uploaded_file(uploaded_file, temp_path: Path) -> bool:
    """Guarda un archivo subido de forma segura."""
    try:
        temp_path.write_bytes(uploaded_file.getvalue())
        return True
    except Exception as e:
        logger.error(f"Error guardando archivo: {e}")
        return False
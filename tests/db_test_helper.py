"""Isolated temporary DB helper using the real persantil schema."""
from __future__ import annotations

import sqlite3
import tempfile
from contextlib import contextmanager
from pathlib import Path
from unittest import mock

import persantil
import insulin_diabetes


@contextmanager
def _connect_row_ctx(db_p: Path):
    conn = sqlite3.connect(db_p, timeout=15)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    try:
        yield conn
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


@contextmanager
def isolated_real_db():
    """Creates an isolated temp directory with a fully initialized Revir Defteri DB."""
    with tempfile.TemporaryDirectory() as tmp:
        tmp_p = Path(tmp)
        db_p = tmp_p / "revir_defteri.db"
        kisi_p = tmp_p / "kisi_verileri"
        kisi_p.mkdir(parents=True, exist_ok=True)
        yedek_p = tmp_p / "oto_yedekler"
        yedek_p.mkdir(parents=True, exist_ok=True)

        patches = [
            mock.patch.object(persantil, "APP_DIR", tmp_p),
            mock.patch.object(persantil, "DB_PATH", db_p),
            mock.patch.object(persantil, "PERSON_DATA_DIR", kisi_p),
            mock.patch.object(persantil, "AUTO_BACKUP_DIR", yedek_p),
            mock.patch.object(persantil, "get_current_db_path", return_value=db_p),
            mock.patch.object(persantil, "get_current_person_data_dir", return_value=kisi_p),
            mock.patch.object(persantil, "connect_master_db", side_effect=lambda: _connect_row_ctx(db_p)),
            mock.patch.object(persantil, "connect_db", side_effect=lambda: _connect_row_ctx(db_p)),
        ]
        for p in patches:
            p.start()
        try:
            try:
                persantil.ensure_db_initialized.clear()
            except Exception:
                pass
            persantil.ensure_db_initialized()
            with _connect_row_ctx(db_p) as c:
                insulin_diabetes.ensure_diabetes_tables(c)
            yield tmp_p, db_p, lambda: _connect_row_ctx(db_p)
        finally:
            try:
                persantil.ensure_db_initialized.clear()
            except Exception:
                pass
            for p in reversed(patches):
                p.stop()

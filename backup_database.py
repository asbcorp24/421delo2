"""Create consistent ZIP backups of the SQLite database."""

import sqlite3
import tempfile
import zipfile
from datetime import datetime
from pathlib import Path


BASE_DIR = Path(__file__).resolve().parent
DB_PATH = BASE_DIR / "app.db"
BACKUP_DIR = BASE_DIR / "backups"


def create_backup(db_path=DB_PATH, backup_dir=BACKUP_DIR) -> Path:
    """Create a consistent SQLite snapshot and pack it into a ZIP archive."""
    db_path = Path(db_path)
    backup_dir = Path(backup_dir)
    if not db_path.is_file():
        raise FileNotFoundError(f"Database file not found: {db_path}")

    backup_dir.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    archive_path = backup_dir / f"otd421_db_{stamp}.zip"

    with tempfile.NamedTemporaryFile(suffix=".db", dir=backup_dir, delete=False) as temp_file:
        snapshot_path = Path(temp_file.name)
    source = None
    snapshot = None
    try:
        source = sqlite3.connect(db_path)
        snapshot = sqlite3.connect(snapshot_path)
        source.backup(snapshot)
        snapshot.commit()
        with zipfile.ZipFile(archive_path, "w", zipfile.ZIP_DEFLATED) as archive:
            archive.write(snapshot_path, arcname=f"app_{stamp}.db")
    finally:
        if snapshot is not None:
            snapshot.close()
        if source is not None:
            source.close()
        snapshot_path.unlink(missing_ok=True)
    return archive_path


def list_backups(backup_dir=BACKUP_DIR):
    backup_dir = Path(backup_dir)
    if not backup_dir.exists():
        return []
    return sorted(
        (path for path in backup_dir.glob("otd421_db_*.zip") if path.is_file()),
        key=lambda path: path.stat().st_mtime,
        reverse=True,
    )


if __name__ == "__main__":
    print(f"Backup created: {create_backup()}")

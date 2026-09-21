import os
from dotenv import load_dotenv

load_dotenv()

SECRET_PASSPHRASE = os.getenv("SECRET_PASSPHRASE")
SUPERADMIN_LOGIN = os.getenv("SUPERADMIN_LOGIN")
SUPERADMIN_PASSWORD = os.getenv("SUPERADMIN_PASSWORD")
if not SECRET_PASSPHRASE:
    raise RuntimeError("SECRET_PASSPHRASE не указана в .env")

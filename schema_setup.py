"""Create/repair the SQLite database for the School Management System."""
from main import ensure_db, DB_PATH
if __name__ == '__main__':
    ensure_db(); print('Database ready:', DB_PATH)

import os
import sys
import subprocess
import time
from git import Repo

REPO_DIR = os.path.dirname(os.path.abspath(__file__))
APP_SCRIPT = "app.py"

def pull_latest_changes():
    try:
        repo = Repo(REPO_DIR)
        origin = repo.remotes.origin
        print("Pulling latest changes from GitHub...")
        origin.pull()
        print("Update completed.")
        return True
    except Exception as e:
        print(f"Ошибка обновления: {e}")
        return False

def restart_app():
    print("Перезапуск приложения...")
    python_executable = sys.executable
    subprocess.Popen([python_executable, APP_SCRIPT])
    print("Приложение запущено. Завершаем updater.")
    sys.exit(0)

if __name__ == "__main__":
    updated = pull_latest_changes()
    if updated:
        # Дадим короткую паузу перед перезапуском
        time.sleep(1)
        restart_app()
    else:
        print("Обновление не выполнено.")

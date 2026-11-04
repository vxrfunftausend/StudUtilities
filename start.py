import sys
import subprocess
import importlib.util
from pathlib import Path


REQUIRED_LIBS = {
    "docx": "python-docx",
}

TARGET_SCRIPT = "docviewer.py"

def is_installed(module_name: str) -> bool:
    return importlib.util.find_spec(module_name) is not None

def install_package(package: str) -> None:
    subprocess.check_call([sys.executable, "-m", "pip", "install", package])

def ensure_dependencies() -> None:
    missing = [pkg for mod, pkg in REQUIRED_LIBS.items() if not is_installed(mod)]
    if not missing:
        return
    for pkg in missing:
        try:
            install_package(pkg)
        except subprocess.CalledProcessError as e:
            sys.exit(1)
    importlib.invalidate_caches()

def launch(script: str) -> None:
    path = Path(__file__).parent / script
    if not path.is_file():
        print(f"Файл '{script}' не найден рядом с main.py")
        sys.exit(1)
    result = subprocess.run([sys.executable, str(path)])
    sys.exit(result.returncode)

if __name__ == "__main__":
    ensure_dependencies()
    launch(TARGET_SCRIPT)

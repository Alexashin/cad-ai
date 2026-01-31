from pathlib import Path
import sys
import pythoncom

# --- bootstrap sys.path so imports work from repo root ---
ROOT = Path(__file__).resolve().parent
# if you moved the KOMPAS helper files into a folder, put its name here:
KOMPAS_SDK_DIR = (
    ROOT / "kompas_sdk"
)  # <- если папки нет, просто создай и положи туда LDefin2D.py etc.

sys.path.insert(0, str(ROOT))
if KOMPAS_SDK_DIR.exists():
    sys.path.insert(0, str(KOMPAS_SDK_DIR))

from cad_ai.ui.app import App  # noqa: E402  (import after sys.path)


def main():
    pythoncom.CoInitialize()
    try:
        app = App()
        app.mainloop()
    finally:
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    main()

import subprocess
import sys


def ensure_dependencies():
    required = ["pandas", "matplotlib", "openpyxl"]
    missing = []
    for pkg in required:
        try:
            __import__(pkg)
        except ImportError:
            missing.append(pkg)
    if missing:
        print(f"Installing missing packages: {', '.join(missing)}")
        subprocess.check_call([sys.executable, "-m", "pip", "install"] + missing)


if __name__ == "__main__":
    ensure_dependencies()
    from gui import ExpenseApp
    app = ExpenseApp()
    app.mainloop()

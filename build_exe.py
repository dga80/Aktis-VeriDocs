import os
import subprocess
import sys

def build():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    os.chdir(script_dir)
    print("Compilando VeriDocs a ejecutable (.exe) con PyInstaller...")

    cmd = [
        sys.executable,
        "-m", "PyInstaller",
        "--noconfirm",
        "--onedir",           # onedir es mucho más rápido al arrancar y más estable en redes
        "--windowed",
        "--name", "VeriDocs",
        "--collect-all", "customtkinter",
        "--collect-all", "winocr",
        "--collect-all", "winrt",
        "--collect-all", "pypdfium2",
        "--add-data", "config.json;.",
        "--add-data", "codigosOCR.xlsx;.",
        "app_gui.py"
    ]

    print("Ejecutando:", " ".join(cmd))
    res = subprocess.run(cmd)
    if res.returncode == 0:
        print("\n¡Compilación exitosa! Ejecutable creado en dist/VeriDocs/VeriDocs.exe")
    else:
        print("\nError durante la compilación.")

if __name__ == "__main__":
    build()

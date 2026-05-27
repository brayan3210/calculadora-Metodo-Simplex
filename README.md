# Simplex Calculator

This repository contains a Python implementation of the Simplex method for solving linear programming problems, along with a **stand‑alone Windows executable** (`simplex_gui.exe`).

## Features
- Graphical user interface built with **Tkinter**.
- Natural‑language model parser.
- Export of the full solution to Excel (`openpyxl`).
- Embedded table rendering for a richer visual experience.
- Custom variable descriptions and detailed result interpretation.

## Quick start (exe)
1. Download the repository (or clone it).
2. Inside the `dist` folder you will find **`simplex_gui.exe`**.
3. Double‑click the exe – no additional Python installation is required.

## Running from source
```bash
# Ensure Python 3.8+ is installed
pip install -r requirements.txt   # only `openpyxl` is needed
python simplex_gui.py
```

## Repository layout
- `simplex_gui.py` – the main application source.
- `dist/simplex_gui.exe` – the packaged executable.
- `README.md` – this file.
- `.gitignore` – excludes build artefacts.
- `requirements.txt` – Python dependencies (currently only `openpyxl`).

## Building the executable yourself
```bash
pip install pyinstaller
pyinstaller --onefile --windowed simplex_gui.py
```
The resulting exe will be placed in `dist/`.

## License
MIT – feel free to use, modify and distribute.

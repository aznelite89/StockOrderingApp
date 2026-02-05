# StockOrderingApp — Run on macOS with `uv`

This project is a simple **Streamlit** app (`app.py`).
The provided `launcher.bat` is for Windows only and simply runs:

```
streamlit run app.py
```

On macOS, the easiest and safest way to run this is with **uv** (no venv, no pip, no Homebrew conflicts).

---

## Prerequisites

- macOS
- Python 3 installed
- `uv` installed

Check:

```bash
uv --version
```

---

## Run the App (no setup required)

run the script directly:-
./run.sh

or

Open Terminal and go to the project folder:

```bash
cd ~/Downloads/StockOrderingApp
```

Run:

```bash
uv run --with streamlit --with pandas --with xlsxwriter streamlit run app.py
```

Your browser will open at:

```
http://localhost:8501
```

---

## If you see `ModuleNotFoundError`

This project has no `requirements.txt`.
If Streamlit reports missing packages, add them like this:

```bash
uv run --with streamlit --with pandas --with requests streamlit run app.py
```

Add any missing package after `--with`.

---

## Optional: Create a one-click launcher

Create a small script so you don’t need to remember the command:

```bash
echo 'uv run --with streamlit streamlit run app.py' > run.sh
chmod +x run.sh
```

Next time, just run:

```bash
./run.sh
```

---

## Why use `uv`?

- No virtualenv needed
- No pip / Homebrew conflicts (PEP 668 safe)
- Fully isolated temporary environment
- Perfect for legacy projects with no documentation

---

## What this project is

- A single-file **Streamlit** Python app
- No backend server
- No database
- Originally designed to be double-clicked on Windows via `.bat`

You are now running it the modern macOS way.

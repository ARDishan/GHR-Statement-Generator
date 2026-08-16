# Payment Schedule Statement Generator

Generates one PDF payment-schedule statement per customer from an Excel file.
The PDF logic lives in `core.py` and is shared by **two interchangeable front
ends** — pick whichever fits:

| | `app_streamlit.py` | `app_tkinter.py` |
|---|---|---|
| Look | Modern browser UI | Native desktop window |
| Run via | `streamlit run app_streamlit.py` | `python app_tkinter.py` |
| Choose output location | Download ZIP, or type a folder path to save to | Native "Browse folder" dialog |
| Package as single .exe | Not typical (Streamlit apps normally stay as `streamlit run`) | Yes, via PyInstaller |

Both run 100% locally on your machine — nothing is uploaded anywhere.

## Folder contents

```
statement_app/
├── core.py              ← shared PDF-generation logic (no UI code)
├── app_streamlit.py      ← browser UI
├── app_tkinter.py         ← desktop UI
├── requirements.txt
├── README.md
└── assets/               ← put your logo/watermark images here
    ├── GHR.png            (branch: GLOBAL HOUSING & REAL ESTATE LTD)
    ├── CED.png              (branch: CORALS EDGE (PVT) LTD)
    └── CPlus.png              (background watermark)
```

If an image is missing from `assets/`, both apps still run — they just skip
that logo/watermark instead of crashing. Any branch name not in the two
listed above falls back to `GHR.png`.

## Setup (once)

```bash
cd statement_app
pip install -r requirements.txt
```

## Option A — Streamlit (recommended, modern UI)

A Streamlit app is technically a small local web server, so it does need to
be *started* each time — but you don't have to type the command yourself.
Use the included double-click launcher instead:

- **Windows:** double-click `run_windows.bat`
- **macOS:** double-click `run_mac.command`
  (first time only: right-click → Open, to bypass Gatekeeper; or run
  `chmod +x run_mac.command` once in Terminal)

Either one runs `launcher.py`, which starts the server in the background,
picks a free port automatically, and opens your default browser to it —
no terminal, no typing. Closing the terminal/console window it opens stops
the server.

If you'd rather run it manually:
```bash
streamlit run app_streamlit.py
```

Flow once it's open: upload the `.xlsx` → preview the data → **Generate
PDFs** → download everything as one ZIP (or download PDFs individually), or
type a folder path and click **Save to folder** to write them straight to
disk on this machine.

### Turning the launcher into a real double-click .exe (no "python" needed)
```bash
pip install pyinstaller
pyinstaller --onefile --name "Statement Generator" launcher.py
```
Then copy `app_streamlit.py`, `core.py`, and `assets/` into the same folder
as the built executable (Streamlit apps are bundled as external files
rather than baked in, so PyInstaller can find and run them). Distribute
that folder — double-clicking the .exe now does everything.

## Option B — Tkinter (native desktop window, packages to a single .exe)

```bash
python app_tkinter.py
```

A window opens with Browse buttons for the Excel file and the output folder,
a Generate button, live progress log, Cancel, and Open Output Folder.

### Package Option B as a standalone .exe (no Python needed to run it)

```bash
pip install pyinstaller
# Windows
pyinstaller --noconsole --onefile --name "Statement Generator" app_tkinter.py
# macOS / Linux
pyinstaller --noconsole --onefile --name "StatementGenerator" app_tkinter.py
```

Copy `assets/` (with your real logo files) next to the built executable in
`dist/`, then distribute that folder. The app looks for `assets/` alongside
the exe at runtime, so you can swap logos later without rebuilding.

## Notes on the Excel format

One row per installment. Rows are grouped first by `CUSTOMER`, then by
`Unit REF ID` within each customer's PDF, so a customer with multiple units
gets one PDF containing a section per unit. The output filename for each
customer's PDF comes from the `FILE NAME` column.

Required columns: BRANCH, PROJECT, CUSTOMER, Unit REF ID, S NO, INSTALLMENT
NO, INSTALLMENT AMT, DUE DATE, PAID AMT, OUTSTANDING, FILE NAME.

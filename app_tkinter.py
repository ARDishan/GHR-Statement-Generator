"""
Payment Schedule Statement Generator — Desktop (Tkinter)

Pick an Excel file, pick an output folder, click Generate.
One PDF is generated per customer.

PDF generation is handled by core.py.

For development:
    python app_tkinter.py

For Windows PyInstaller build:
    pyinstaller StatementGenerator.spec
"""

import os
import sys
import queue
import threading
import traceback
from datetime import datetime

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import core


# ---------------------------------------------------------------------------
# Runtime paths
# ---------------------------------------------------------------------------

ASSETS_DIR = core.ASSETS_DIR


class App(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("Payment Schedule Statement Generator")
        self.geometry("720x520")
        self.minsize(640, 460)

        self.excel_path = tk.StringVar()
        self.output_dir = tk.StringVar()
        self.status_var = tk.StringVar(value="Ready.")

        # Communication between worker thread and Tkinter main thread.
        self.ui_queue = queue.Queue()

        self.stop_flag = threading.Event()
        self.worker_thread = None
        self.last_output_dir = None

        self._build_ui()

        # All UI updates are processed by the Tkinter main thread.
        self.after(100, self._poll_ui_queue)

        self.protocol("WM_DELETE_WINDOW", self._on_close)

        self._check_assets()

    # -----------------------------------------------------------------------
    # UI
    # -----------------------------------------------------------------------

    def _build_ui(self):
        pad = {"padx": 10, "pady": 6}

        frm_top = ttk.Frame(self)
        frm_top.pack(fill="x", **pad)

        ttk.Label(
            frm_top,
            text="Excel file (.xlsx):"
        ).grid(row=0, column=0, sticky="w")

        ttk.Entry(
            frm_top,
            textvariable=self.excel_path
        ).grid(
            row=0,
            column=1,
            sticky="ew",
            padx=6
        )

        ttk.Button(
            frm_top,
            text="Browse...",
            command=self._pick_excel
        ).grid(row=0, column=2)

        ttk.Label(
            frm_top,
            text="Output folder:"
        ).grid(
            row=1,
            column=0,
            sticky="w",
            pady=(8, 0)
        )

        ttk.Entry(
            frm_top,
            textvariable=self.output_dir
        ).grid(
            row=1,
            column=1,
            sticky="ew",
            padx=6,
            pady=(8, 0)
        )

        ttk.Button(
            frm_top,
            text="Browse...",
            command=self._pick_output_dir
        ).grid(
            row=1,
            column=2,
            pady=(8, 0)
        )

        frm_top.columnconfigure(1, weight=1)

        frm_btns = ttk.Frame(self)
        frm_btns.pack(fill="x", **pad)

        self.generate_btn = ttk.Button(
            frm_btns,
            text="Generate PDFs",
            command=self._start_generate
        )
        self.generate_btn.pack(side="left")

        self.cancel_btn = ttk.Button(
            frm_btns,
            text="Cancel",
            command=self._cancel,
            state="disabled"
        )
        self.cancel_btn.pack(side="left", padx=8)

        self.open_folder_btn = ttk.Button(
            frm_btns,
            text="Open Output Folder",
            command=self._open_output_folder,
            state="disabled"
        )
        self.open_folder_btn.pack(side="left")

        self.progress = ttk.Progressbar(
            self,
            mode="indeterminate"
        )
        self.progress.pack(
            fill="x",
            padx=10,
            pady=(0, 6)
        )

        frm_log = ttk.LabelFrame(
            self,
            text="Log"
        )
        frm_log.pack(
            fill="both",
            expand=True,
            padx=10,
            pady=(0, 6)
        )

        self.log_text = tk.Text(
            frm_log,
            wrap="word",
            state="disabled",
            height=16
        )
        self.log_text.pack(
            side="left",
            fill="both",
            expand=True
        )

        scrollbar = ttk.Scrollbar(
            frm_log,
            command=self.log_text.yview
        )
        scrollbar.pack(
            side="right",
            fill="y"
        )

        self.log_text.configure(
            yscrollcommand=scrollbar.set
        )

        status_bar = ttk.Label(
            self,
            textvariable=self.status_var,
            anchor="w",
            relief="sunken"
        )
        status_bar.pack(
            fill="x",
            side="bottom"
        )

    # -----------------------------------------------------------------------
    # Asset check
    # -----------------------------------------------------------------------

    def _check_assets(self):
        required_images = (
            "GHR.png",
            "CED.png",
        )

        optional_watermarks = (
            "EBIZ.png",
            "CPlus.png",
        )

        missing = []

        for filename in required_images:
            if not os.path.exists(
                os.path.join(ASSETS_DIR, filename)
            ):
                missing.append(filename)

        if not any(
            os.path.exists(
                os.path.join(ASSETS_DIR, filename)
            )
            for filename in optional_watermarks
        ):
            missing.append("EBIZ.png or CPlus.png")

        if missing:
            self._log(
                "Warning: missing asset(s): "
                + ", ".join(missing)
            )

        # Font check is already performed by core.py during import.
        self._log(
            f"Assets directory: {ASSETS_DIR}"
        )

    # -----------------------------------------------------------------------
    # File dialogs
    # -----------------------------------------------------------------------

    def _pick_excel(self):
        path = filedialog.askopenfilename(
            title="Select Excel file",
            filetypes=[
                ("Excel files", "*.xlsx *.xls"),
                ("All files", "*.*"),
            ],
        )

        if path:
            self.excel_path.set(path)

    def _pick_output_dir(self):
        path = filedialog.askdirectory(
            title="Select output folder"
        )

        if path:
            self.output_dir.set(path)

    # -----------------------------------------------------------------------
    # Logging
    # -----------------------------------------------------------------------

    def _log(self, message):
        timestamp = datetime.now().strftime("%H:%M:%S")

        self.ui_queue.put(
            ("log", f"[{timestamp}] {message}")
        )

    def _append_log(self, message):
        self.log_text.configure(state="normal")
        self.log_text.insert("end", message + "\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    # -----------------------------------------------------------------------
    # UI queue
    # -----------------------------------------------------------------------

    def _poll_ui_queue(self):
        try:
            while True:
                event_type, payload = self.ui_queue.get_nowait()

                if event_type == "log":
                    self._append_log(payload)

                elif event_type == "done":
                    self._handle_done(payload)

                elif event_type == "cancelled":
                    self._handle_cancelled(payload)

                elif event_type == "error":
                    self._handle_error(payload)

        except queue.Empty:
            pass

        if self.winfo_exists():
            self.after(100, self._poll_ui_queue)

    # -----------------------------------------------------------------------
    # Running state
    # -----------------------------------------------------------------------

    def _set_running(self, running):
        if running:
            self.generate_btn.configure(state="disabled")
            self.cancel_btn.configure(state="normal")
            self.open_folder_btn.configure(state="disabled")
            self.progress.start(10)
        else:
            self.generate_btn.configure(state="normal")
            self.cancel_btn.configure(state="disabled")
            self.progress.stop()

    # -----------------------------------------------------------------------
    # Start generation
    # -----------------------------------------------------------------------

    def _start_generate(self):
        if (
            self.worker_thread is not None
            and self.worker_thread.is_alive()
        ):
            return

        excel_path = self.excel_path.get().strip()
        output_dir = self.output_dir.get().strip()

        if not excel_path or not os.path.isfile(excel_path):
            messagebox.showerror(
                "Missing file",
                "Please select a valid Excel file."
            )
            return

        if not output_dir:
            messagebox.showerror(
                "Missing folder",
                "Please select an output folder."
            )
            return

        self.stop_flag.clear()
        self.last_output_dir = None

        self.status_var.set("Generating...")
        self._set_running(True)

        self._append_log("")
        self._append_log(
            "------------------------------------------------------------"
        )
        self._append_log(
            f"Started: {os.path.basename(excel_path)}"
        )
        self._append_log(
            f"Output: {output_dir}"
        )

        self.worker_thread = threading.Thread(
            target=self._run_worker,
            args=(excel_path, output_dir),
            daemon=True,
            name="PDFGeneratorWorker",
        )

        self.worker_thread.start()

    # -----------------------------------------------------------------------
    # Worker thread
    # -----------------------------------------------------------------------

    def _run_worker(self, excel_path, output_dir):
        try:
            files = core.generate_pdfs(
                excel_path,
                output_dir,
                self._log,
                self.stop_flag
            )

            # Do not touch Tkinter widgets from this worker thread.
            if self.stop_flag.is_set():
                self.ui_queue.put(
                    ("cancelled", {
                        "count": len(files),
                        "output_dir": output_dir,
                    })
                )
            else:
                self.ui_queue.put(
                    ("done", {
                        "count": len(files),
                        "output_dir": output_dir,
                    })
                )

        except Exception as exc:
            error_text = (
                f"{exc}\n\n"
                f"{traceback.format_exc()}"
            )

            self.ui_queue.put(
                ("error", {
                    "message": str(exc),
                    "traceback": error_text,
                })
            )

    # -----------------------------------------------------------------------
    # Worker results — main Tkinter thread
    # -----------------------------------------------------------------------

    def _handle_done(self, payload):
        count = payload["count"]
        output_dir = payload["output_dir"]

        self.last_output_dir = output_dir

        self._set_running(False)

        self.status_var.set(
            f"Done — {count} PDF(s) generated."
        )

        self._append_log(
            f"Done. {count} PDF(s) written to '{output_dir}'."
        )

        self.open_folder_btn.configure(
            state="normal"
        )

        messagebox.showinfo(
            "Generation complete",
            f"{count} PDF(s) generated successfully."
        )

    def _handle_cancelled(self, payload):
        count = payload["count"]
        output_dir = payload["output_dir"]

        self.last_output_dir = output_dir

        self._set_running(False)

        self.status_var.set(
            "Cancelled."
        )

        self._append_log(
            f"Stopped. {count} PDF(s) written before cancellation."
        )

        if count > 0 and os.path.isdir(output_dir):
            self.open_folder_btn.configure(
                state="normal"
            )

    def _handle_error(self, payload):
        self._set_running(False)

        self.status_var.set(
            "Failed — see log."
        )

        self._append_log(
            f"ERROR: {payload['message']}"
        )

        self._append_log(
            payload["traceback"]
        )

        messagebox.showerror(
            "Generation failed",
            payload["message"]
        )

    # -----------------------------------------------------------------------
    # Cancel
    # -----------------------------------------------------------------------

    def _cancel(self):
        if (
            self.worker_thread is None
            or not self.worker_thread.is_alive()
        ):
            return

        self.stop_flag.set()

        self.status_var.set(
            "Cancelling..."
        )

        self._append_log(
            "Cancellation requested. "
            "The current PDF will finish before stopping."
        )

    # -----------------------------------------------------------------------
    # Open output folder
    # -----------------------------------------------------------------------

    def _open_output_folder(self):
        target = (
            self.last_output_dir
            or self.output_dir.get().strip()
        )

        if not target or not os.path.isdir(target):
            return

        try:
            if sys.platform.startswith("win"):
                os.startfile(target)

            elif sys.platform == "darwin":
                os.system(
                    f'open "{target}"'
                )

            else:
                os.system(
                    f'xdg-open "{target}"'
                )

        except Exception as exc:
            messagebox.showerror(
                "Unable to open folder",
                str(exc)
            )

    # -----------------------------------------------------------------------
    # Window close
    # -----------------------------------------------------------------------

    def _on_close(self):
        if (
            self.worker_thread is not None
            and self.worker_thread.is_alive()
        ):
            answer = messagebox.askyesno(
                "Generation in progress",
                "PDF generation is still running.\n\n"
                "Do you want to cancel it and close the application?"
            )

            if not answer:
                return

            self.stop_flag.set()

        self.destroy()


# ---------------------------------------------------------------------------
# Application entry point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    app = App()
    app.mainloop()
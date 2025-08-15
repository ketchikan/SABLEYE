import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
import pandas as pd
from utils.email_sender import OutlookEmailSender
from utils.pathing import resource_path
import threading

class App(tk.Tk):
    """
    Simple multi-page Tkinter app:
      Home -> choose "Disputes" or "Pay Reports"
      Each choice shows a dedicated page you can build out later.
    """
    def __init__(self):
        super().__init__()
        self.title("Email Automation")
        self.geometry("1080x720")

        # SETTING APP ICONS
        root = self
        try:
            root.iconbitmap(resource_path("assets", "app_icon.ico"))
        except Exception:
            # Fallback for non-ICO
            try:
                icon = tk.PhotoImage(file=resource_path("assets", "header.png"))
                root.iconphoto(True, icon)
            except Exception:
                pass

        # Root container that holds all pages stacked on top of each other
        container = ttk.Frame(self, padding=12)
        container.pack(fill="both", expand=True)

        # Make container stretch
        container.rowconfigure(0, weight=1)
        container.columnconfigure(0, weight=1)

        # Register and create pages
        self.pages = {}
        for Page in (HomePage, DisputesPage, PayReportsPage):
            page = Page(parent=container, controller=self)
            self.pages[Page.__name__] = page
            page.grid(row=0, column=0, sticky="nsew")

        self.show_page("HomePage")

    def show_page(self, name: str):
        """Raise a page by its class name stored in self.pages."""
        frame = self.pages[name]
        frame.tkraise()


class HomePage(ttk.Frame):
    def __init__(self, parent, controller: App):
        super().__init__(parent)
        self.controller = controller

        # Layout
        self.columnconfigure(0, weight=1)

        title = ttk.Label(self, text="Email Automation", font=("Segoe UI", 18, "bold"))
        title.grid(row=0, column=0, pady=(10, 20))

        desc = ttk.Label(
            self,
            text="Choose what you want to work on:",
            font=("Segoe UI", 11)
        )
        desc.grid(row=1, column=0, pady=(0, 20))

        # Buttons row
        btns = ttk.Frame(self)
        btns.grid(row=2, column=0, pady=10)

        disputes_btn = ttk.Button(
            btns, text="Disputes",
            command=lambda: controller.show_page("DisputesPage")
        )
        disputes_btn.grid(row=0, column=0, padx=8, ipadx=12, ipady=6)

        pay_reports_btn = ttk.Button(
            btns, text="Pay Reports",
            command=lambda: controller.show_page("PayReportsPage")
        )
        pay_reports_btn.grid(row=0, column=1, padx=8, ipadx=12, ipady=6)


class DisputesPage(ttk.Frame):
    EXPECTED_SHEET = "Emails"
    # Columns C:N inclusive (12 columns), in order:
    EXPECTED_COLUMNS = [
        "Submitter",
        "Project ID",
        "Appt Date",
        "Requested Outcome",
        "Context",
        "Closer",
        "Outcome",
        "Outcome Note",
        "Closer Manager",
        "Setter Mgr First",
        "Closer Mgr First",
        "Email-To",
    ]

    def __init__(self, parent, controller: App):
        super().__init__(parent)
        self.controller = controller

        # ----- Setup -----
        self.preview_var = tk.BooleanVar(value=False)
        self.on_behalf_var = tk.StringVar(value="disputes@blueravensolar.com")

        # ----- State -----
        self.selected_file = tk.StringVar(value="")
        self.df = None

        # ----- Layout -----
        self.columnconfigure(0, weight=1)

        # Row 0: Header
        header = ttk.Label(self, text="Disputes Email Page", font=("Segoe UI", 16, "bold"))
        header.grid(row=0, column=0, pady=(10, 12))

        # Row 1: Info
        info = ttk.Label(
            self,
            text="Use this page to send the results of the Disputes process.\n"
                 "Please keep in mind that you must have access to the inbox disputes@blueravensolar.com before sending these emails.",
            justify="center"
        )
        info.grid(row=1, column=0, pady=(0,20))

        # Row 2: Subheading
        sub = ttk.Label(self, text="Select the Excel file that contains the Disputes results.")
        sub.grid(row=2, column=0, pady=(0,10))

        # Row 3: File Picker (tighter vertical spacing)
        row_file = ttk.Frame(self)
        row_file.grid(row=3, column=0, sticky="ew", pady=(0, 2))  # was (0, 8)
        row_file.columnconfigure(1, weight=1)

        ttk.Label(row_file, text="Source file:").grid(row=0, column=0, padx=(0, 6))
        self.file_entry = ttk.Entry(row_file, textvariable=self.selected_file)
        self.file_entry.grid(row=0, column=1, sticky="ew")
        ttk.Button(row_file, text="Browse…", command=self.on_browse).grid(row=0, column=2, padx=(6, 0))

        # Row 4: Actions (tighter)
        row_actions = ttk.Frame(self)
        row_actions.grid(row=4, column=0, sticky="w", pady=(0, 2))  # was (0, 5)
        ttk.Button(row_actions, text="Validate file", command=self.on_validate).grid(row=0, column=0, padx=(0, 6))
        ttk.Button(row_actions, text="Load data", command=self.on_load_data).grid(row=0, column=1)
        ttk.Button(row_actions, text="Send All Emails", command=self.on_send_emails).grid(row=0, column=2)

        # Row 5: Output / status box
        self._build_output_tabs()

        # Row 6: Back Button
        back = ttk.Button(self, text="← Back to Home", command=lambda: controller.show_page("HomePage"))
        back.grid(row=6, column=0, pady=(20, 0))


        # ---------- handlers ----------
    def on_browse(self):
        path = filedialog.askopenfilename(
            title="Select disputes source file",
            filetypes=[
                ("Excel files", "*.xlsx *.xlsm"),
                ("All files", "*.*"),
            ],
        )
        if path:
            self.selected_file.set(path)
            self.log(f"Selected file: {path}")

    def on_validate(self):
        path_str = self.selected_file.get().strip()
        if not path_str:
            messagebox.showerror("Missing file", "Please choose a source file first.")
            return

        p = Path(path_str)
        if not p.exists():
            messagebox.showerror("Not found", f"File does not exist:\n{p}")
            return

        # Try reading workbook metadata to confirm it’s an Excel file we can open.
        try:
            # Lightweight check: list sheet names (no full read of all sheets)
            xls = pd.ExcelFile(p)
            self.log(f"✅ File is readable. Sheets: {', '.join(xls.sheet_names)}")
        except PermissionError:
            self.log(f"❌ Permission error. Close the file if it’s open in Excel:\n{p}")
            messagebox.showerror("File locked", "Close the file in Excel and try again.")
        except Exception as e:
            self.log(f"❌ Could not read file:\n{e}")
            messagebox.showerror("Invalid file", f"Unable to open as Excel:\n{e}")

    def on_load_data(self):
        """Load C:N from the 'Emails' sheet, validate headers, and store DataFrame."""
        path = self._require_path()
        if not path:
            return

        try:
            # First confirm the sheet exists
            xls = pd.ExcelFile(path)
            if self.EXPECTED_SHEET not in xls.sheet_names:
                messagebox.showerror(
                    "Missing sheet",
                    f"Sheet '{self.EXPECTED_SHEET}' was not found.\nSheets available: {', '.join(xls.sheet_names)}",
                )
                self.log(f"❌ Missing sheet '{self.EXPECTED_SHEET}'.")
                return

            # Read only columns C:N (12 columns)
            df = pd.read_excel(
                path,
                sheet_name=self.EXPECTED_SHEET,
                usecols="C:N",
                engine="openpyxl",
            )

            # Normalize header whitespace
            df.columns = [str(c).strip() for c in df.columns]

            # Validate exact schema in order
            got = list(df.columns)
            exp = list(self.EXPECTED_COLUMNS)
            if got != exp:
                # Build a clear diff message
                problems = []
                for i, (g, e) in enumerate(zip(got, exp), start=1):
                    if g != e:
                        problems.append(f"  Col {i}: expected '{e}', got '{g}'")
                # Handle length mismatch
                if len(got) != len(exp):
                    problems.append(f"  Column count mismatch: expected {len(exp)}, got {len(got)}")
                msg = "Header mismatch:\n" + "\n".join(problems)
                messagebox.showerror("Unexpected columns", msg)
                self.log(f"❌ {msg}")
                return

            # Store and report
            self.df = df
            self.log(f"✅ Loaded {len(df):,} rows from '{self.EXPECTED_SHEET}' (C:N).")
            self._render_preview_df(self.df, limit=50)
            self.tabs.select(self.preview_frame)  # switch to the Preview tab

        except PermissionError:
            self.log(f"❌ Permission error. Close the file if it’s open in Excel:\n{path}")
            messagebox.showerror("File locked", "Close the file in Excel and try again.")
        except Exception as e:
            self.log(f"❌ Load failed: {e}")
            messagebox.showerror("Load failed", str(e))
    
    def on_send_emails(self):
        """Validate and kick off a background send."""
        if self.df is None or self.df.empty:
            messagebox.showerror("No data", "Load data from the 'Emails' sheet first.")
            return

        # Basic sanity check for required column
        if "Email-To" not in self.df.columns:
            messagebox.showerror("Missing column", "Expected 'Email-To' column not found.")
            return

        self.log("\n[Sending…]\n")
        self._set_busy(True)

        thread = threading.Thread(target=self._send_worker, daemon=True)
        thread.start()

    def _send_worker(self):
        """Background thread: iterate rows and send/preview each email."""
        try:
            preview = self.preview_var.get()
            send_on_behalf = self.on_behalf_var.get().strip() or None

            sent = 0
            skipped = 0

            with OutlookEmailSender(send_on_behalf_of=send_on_behalf, preview=preview) as sender:
                for idx, row in self.df.iterrows():
                    # to_addr = str(row.get("Email-To") or "").strip()
                    to_addr = "jacob.r.west@sunpower.com"
                    if not to_addr:
                        skipped += 1
                        self._ui_log(f"– Skipped row {idx}: missing Email-To")
                        continue

                    # Build a simple subject + HTML body.
                    # Adjust formatting to your needs or plug in a Jinja renderer later.
                    project_id = row.get("Project ID", "")
                    requested = row.get("Requested Outcome", "")
                    outcome = row.get("Outcome", "")
                    outcome_note = row.get("Outcome Note", "")
                    submitter = row.get("Submitter", "")
                    appt_date = row.get("Appt Date", "")

                    subject = f"Dispute Result – Project {project_id} – {requested or outcome}"
                    html_body = f"""
                    <html>
                    <body style="font-family:Segoe UI, Arial, sans-serif; font-size:12pt;">
                        <p>Hi,</p>
                        <p>The dispute result for <b>Project {project_id}</b> is below.</p>
                        <table cellpadding="6" cellspacing="0" border="0" style="border-collapse:collapse;">
                        <tr><td><b>Submitter</b></td><td>{submitter}</td></tr>
                        <tr><td><b>Appt Date</b></td><td>{appt_date}</td></tr>
                        <tr><td><b>Requested Outcome</b></td><td>{requested}</td></tr>
                        <tr><td><b>Final Outcome</b></td><td>{outcome}</td></tr>
                        <tr><td><b>Note</b></td><td>{outcome_note}</td></tr>
                        </table>
                        <p style="margin-top:14px;">Regards,<br>Disputes Team</p>
                    </body>
                    </html>
                    """

                    try:
                        sender.send_html(
                            html_body=html_body,
                            to=to_addr,
                            subject=subject,
                            # Optionally add cc/bcc here if you have columns for them
                            # cc=row.get("CC", ""),
                            # bcc=row.get("BCC", ""),
                            # reply_to="disputes@sunpower.com",
                            preview=preview,  # honors UI toggle
                            send_on_behalf_of=send_on_behalf  # per-message override (optional)
                        )
                        sent += 1
                        self._ui_log(f"✓ Queued row {idx} → {to_addr}")
                    except Exception as e:
                        self._ui_log(f"✗ Failed row {idx} → {to_addr}: {e}")

            self._ui_log(f"\nDone. Sent/Previewed: {sent} | Skipped: {skipped}\n")

        except Exception as e:
            self._ui_log(f"❌ Error: {e}\n")
        finally:
            self._set_busy(False)

    def _ui_log(self, text: str):
        self.after(0, lambda: (self.output.insert("end", text + "\n"),
                            self.output.see("end")))

    def _set_busy(self, busy: bool):
        # You can disable buttons/entries here while sending
        state = "disabled" if busy else "normal"
        for w in (self.file_entry,):
            try:
                w.config(state=state)
            except Exception:
                pass

    # ---------- helpers ----------
    def _require_path(self) -> Path | None:
        path_str = self.selected_file.get().strip()
        if not path_str:
            messagebox.showerror("Missing file", "Please choose a source file first.")
            return None
        p = Path(path_str)
        if not p.exists():
            messagebox.showerror("Not found", f"File does not exist:\n{p}")
            return None
        return p
    
    def _build_output_tabs(self):
        """Create a tabbed area with Log and Preview tables."""
        # Notebook
        self.tabs = ttk.Notebook(self)
        self.tabs.grid(row=5, column=0, sticky="nsew")
        self.rowconfigure(5, weight=1)

        # Log tab
        self.log_frame = ttk.Frame(self.tabs)
        self.tabs.add(self.log_frame, text="Log")

        self.output = tk.Text(self.log_frame, height=10, wrap="word", borderwidth=1, highlightthickness=0)
        self.output.pack(fill="both", expand=True)

        # Preview tab
        self.preview_frame = ttk.Frame(self.tabs)
        self.tabs.add(self.preview_frame, text="Preview")

        # Treeview + scrollbars
        self.preview_tree = ttk.Treeview(self.preview_frame, show="headings")
        yscroll = ttk.Scrollbar(self.preview_frame, orient="vertical", command=self.preview_tree.yview)
        xscroll = ttk.Scrollbar(self.preview_frame, orient="horizontal", command=self.preview_tree.xview)
        self.preview_tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)

        self.preview_tree.grid(row=0, column=0, sticky="nsew")
        yscroll.grid(row=0, column=1, sticky="ns")
        xscroll.grid(row=1, column=0, sticky="ew")

        self.preview_frame.rowconfigure(0, weight=1)
        self.preview_frame.columnconfigure(0, weight=1)

    def _render_preview_df(self, df: pd.DataFrame, limit: int = 50):
        """Render a DataFrame to the Treeview."""
        # Clear existing
        for col in self.preview_tree["columns"]:
            self.preview_tree.heading(col, text="")
        self.preview_tree.delete(*self.preview_tree.get_children())

        # Choose columns and format datelike columns
        view_df = df.head(limit).copy()

        # Try to pretty-format dates
        # for col in view_df.columns:
        #     if pd.api.types.is_datetime64_any_dtype(view_df[col]):
        #         view_df[col] = view_df[col].dt.strftime("%m/%d/%Y")
        #     else:
        #         # If it's object and looks like datelike, try coercion
        #         try:
        #             coerced = pd.to_datetime(view_df[col], errors="coerce")
        #             if coerced.notna().any():
        #                 view_df[col] = coerced.dt.strftime("%m/%d/%Y").where(coerced.notna(), view_df[col])
        #         except Exception:
        #             pass

        cols = list(view_df.columns)
        self.preview_tree["columns"] = cols

        # Headings
        for c in cols:
            self.preview_tree.heading(c, text=c)
            self.preview_tree.column(c, width=100, anchor="w")  # initial width

        # Rows
        for _, row in view_df.iterrows():
            values = ["" if pd.isna(v) else str(v) for v in row.tolist()]
            self.preview_tree.insert("", "end", values=values)

        # Auto-size (approx) based on content
        self.after(0, lambda: self._autosize_columns(cols))

    def _autosize_columns(self, cols):
        """Rough auto-size: widen columns based on header+sample content."""
        # Measure a few rows (Treeview doesn’t have text measurement; we approximate by char count)
        sample_count = min(25, len(self.preview_tree.get_children()))
        per_col_max = {c: len(str(c)) for c in cols}

        for iid in self.preview_tree.get_children()[:sample_count]:
            vals = self.preview_tree.item(iid, "values")
            for c, v in zip(cols, vals):
                per_col_max[c] = max(per_col_max[c], len(str(v)))

        # Convert char count to pixels (approx 7 px per char; tweak as needed)
        for c in cols:
            px = min(600, max(60, per_col_max[c] * 7))
            self.preview_tree.column(c, width=px)

    
    def log(self, text: str):
        self.output.insert("end", text + "\n")
        self.output.see("end")


class PayReportsPage(ttk.Frame):
    def __init__(self, parent, controller: App):
        super().__init__(parent)
        self.controller = controller

        self.columnconfigure(0, weight=1)

        header = ttk.Label(self, text="Pay Reports Email Page", font=("Segoe UI", 16, "bold"))
        header.grid(row=0, column=0, pady=(10, 12))

        info = ttk.Label(
            self,
            text="This is where you’ll build the Pay Reports workflow.\n"
                 "Add controls here for choosing the pay period file, target, and template.",
            justify="center"
        )
        info.grid(row=1, column=0, pady=(0, 20))

        # Example placeholder controls
        sample_frame = ttk.LabelFrame(self, text="Next Steps")
        sample_frame.grid(row=2, column=0, padx=12, pady=12, sticky="n")
        ttk.Label(sample_frame, text="• Pick SharePoint file\n• Build context\n• Render & Send").pack(padx=10, pady=10)

        back = ttk.Button(self, text="← Back to Home", command=lambda: controller.show_page("HomePage"))
        back.grid(row=3, column=0, pady=(20, 0))


if __name__ == "__main__":
    App().mainloop()

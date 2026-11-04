"""Desktop dashboard for the same report pipeline used by the CLI."""

from copy import deepcopy
import os
from pathlib import Path
from queue import Empty, Queue
import sys
from tempfile import TemporaryDirectory
from threading import Event, Thread
import tkinter as tk
from tkinter import font as tkfont
from tkinter import filedialog, messagebox, ttk
import webbrowser

from .actions import workspace_identity
from .config import DEFAULTS, load_config, load_preferences, save_preferences
from .organizations import OrganizationCache, load_settings
from .pipeline import report
from .results import document_row, read_results
from .workbook import read_source
from .site import WindowsReader
from .reviewers import parse_names
from .filtering import normalized

CARDS = (
    ("Qualified", "Qualified"),
    ("Yes", "Yes"),
    ("No", "No"),
    ("Unresolved", "Unresolved Actions"),
    ("Excluded", "Excluded"),
)
DOCUMENT_COLUMNS = (
    ("result", "Flawless", 80),
    ("document", "Document ID", 95),
    ("manager", "Author Manager", 170),
    ("organization", "Organization", 240),
    ("phase", "Phase", 100),
    ("title", "Document", 300),
    ("submitted", "Submitted", 95),
    ("url", "Document WS URL", 420),
    ("notes", "Notes", 200),
)
RESULT_FILTERS = ("All", "Yes", "No", "Unresolved")
ROW_TAGS = {"Yes": "yes", "No": "no"}


class Dashboard:
    def __init__(self, root):
        self.root = root
        self.events = Queue()
        self.running = False
        self.cancel = Event()
        self.manager_items = {}
        self.documents_data = []
        self.report_bytes = None
        self.run_month = ""
        self.organization_refresh = False
        self.organization_warning = ""
        self.filter_popup = None
        self.filter_bindings = []
        self.organization_error = ""
        try:
            load_settings()
            OrganizationCache()
        except (OSError, ValueError) as exc:
            self.organization_error = str(exc)
            messagebox.showerror(
                "Organization settings/cache error", str(exc), parent=root
            )
        root.title("xCCB Flawless Submission Dashboard")
        self.scale = max(1.0, root.winfo_fpixels("1i") / 96)
        root.geometry(f"{self.px(1180)}x{self.px(740)}")
        root.minsize(self.px(800), self.px(600))
        self.config = deepcopy(DEFAULTS)
        home = (
            Path(sys.executable).parent if getattr(sys, "frozen", False) else Path.cwd()
        )
        config_path = home / "config.json"
        if config_path.exists():
            try:
                self.config = load_config(config_path)
            except Exception as exc:
                messagebox.showerror("Configuration error", str(exc), parent=root)
        try:
            self.preferences = load_preferences(home / "reviewer_names.txt")
        except (OSError, ValueError) as exc:
            self.preferences = {}
            self.organization_error = str(exc)
            messagebox.showerror("Settings error", str(exc), parent=root)
        self.configure_styles()

        frame = ttk.Frame(root, padding=12)
        frame.pack(fill="both", expand=True)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(3, weight=1)

        header = ttk.Frame(frame)
        header.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        header.columnconfigure(0, weight=1)
        ttk.Label(
            header, text="xCCB Flawless Submission Dashboard", style="Title.TLabel"
        ).grid(row=0, column=0, sticky="w")
        ttk.Label(
            header,
            text="Admin forum · listed or blank reviewers · Phase excludes WLA / LSI modification · "
            "Actions checked live on xCCB",
            style="Muted.TLabel",
        ).grid(row=1, column=0, sticky="w")

        self.source = tk.StringVar()
        self.evidence = home / "outputs" / "evidence"
        toolbar = ttk.Frame(frame)
        toolbar.grid(row=1, column=0, sticky="ew", pady=(0, 8))
        toolbar.columnconfigure(1, weight=1)
        ttk.Button(toolbar, text="Upload", command=self.choose_source).grid(
            row=0, column=0, padx=(0, 6)
        )
        ttk.Entry(toolbar, textvariable=self.source, width=20).grid(
            row=0, column=1, sticky="ew"
        )
        self.run_button = ttk.Button(
            toolbar, text="Process", command=self.start, style="Accent.TButton"
        )
        self.run_button.grid(row=0, column=4, padx=(8, 0))
        self.stop_button = ttk.Button(
            toolbar, text="Force stop", command=self.stop, state="disabled"
        )
        self.stop_button.grid(row=0, column=5, padx=(6, 0))
        self.export_button = ttk.Button(
            toolbar, text="Export Excel…", command=self.export, state="disabled"
        )
        self.export_button.grid(row=0, column=6, padx=(6, 0))

        body = ttk.Frame(frame)
        body.grid(row=3, column=0, sticky="nsew")
        body.columnconfigure(1, weight=1)
        body.rowconfigure(0, weight=1)
        self.build_names_panel(body)
        main = ttk.Frame(body)
        main.grid(row=0, column=1, sticky="nsew", padx=(10, 0))
        main.columnconfigure(0, weight=1)
        main.rowconfigure(1, weight=1)
        cards = ttk.Frame(main)
        cards.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        self.card_values = {}
        for column, (label, key) in enumerate(CARDS):
            cards.columnconfigure(column, weight=1, uniform="cards")
            card = ttk.Frame(cards, style="Card.TFrame", padding=(10, 4))
            card.grid(
                row=0, column=column, sticky="ew", padx=(0 if column == 0 else 6, 0)
            )
            self.card_values[key] = tk.StringVar(value="–")
            ttk.Label(
                card, textvariable=self.card_values[key], style=f"Card{label}.TLabel"
            ).pack(anchor="w")
            ttk.Label(card, text=label, style="CardLabel.TLabel").pack(anchor="w")

        self.tabs = ttk.Notebook(main)
        self.tabs.grid(row=1, column=0, sticky="nsew")
        self.build_documents_tab()
        self.build_managers_tab()

        footer = ttk.Frame(frame)
        footer.grid(row=4, column=0, sticky="ew", pady=(8, 0))
        footer.columnconfigure(0, weight=1)
        self.status = tk.StringVar(
            value="Upload a workbook, check the listed names, then press Process."
        )
        ttk.Label(footer, textvariable=self.status, style="Muted.TLabel").grid(
            row=0, column=0, sticky="w"
        )
        self.progress = ttk.Progressbar(
            footer, mode="indeterminate", length=self.px(140)
        )
        self.progress.grid(row=0, column=1, sticky="e")
        self.progress.grid_remove()
        self.run_indicator = ttk.Label(footer, text="Stopped", style="Stopped.TLabel")
        self.run_indicator.grid(row=0, column=2, sticky="e", padx=(12, 0))
        root.protocol("WM_DELETE_WINDOW", self.close)
        root.after(100, self.poll)

    def px(self, value):
        return round(value * self.scale)

    def configure_styles(self):
        style = ttk.Style(self.root)
        style.configure("Title.TLabel", font=("Segoe UI Semibold", 15))
        style.configure("Muted.TLabel", foreground="#5F6B7A")
        style.configure(
            "Running.TLabel", foreground="#1E7B34", font=("Segoe UI Semibold", 9)
        )
        style.configure(
            "Stopped.TLabel", foreground="#5F6B7A", font=("Segoe UI Semibold", 9)
        )
        style.configure("Hint.TLabel", foreground="#5F6B7A", font=("Segoe UI", 8))
        style.configure("Accent.TButton", font=("Segoe UI Semibold", 9))
        style.configure(
            "Card.TFrame", background="#FFFFFF", relief="solid", borderwidth=1
        )
        style.configure("CardLabel.TLabel", background="#FFFFFF", foreground="#5F6B7A")
        for label, colour in (
            ("Qualified", "#1F3B5C"),
            ("Yes", "#1E7B34"),
            ("No", "#B3261E"),
            ("Unresolved", "#A15C00"),
            ("Excluded", "#5F6B7A"),
        ):
            style.configure(
                f"Card{label}.TLabel",
                background="#FFFFFF",
                foreground=colour,
                font=("Segoe UI Semibold", 15),
            )
        style.configure("Treeview", rowheight=self.px(24))
        style.configure("Treeview.Heading", font=("Segoe UI Semibold", 9))

    def build_names_panel(self, parent):
        panel = ttk.LabelFrame(parent, text="Listed names", padding=8)
        panel.grid(row=0, column=0, sticky="ns")
        panel.rowconfigure(1, weight=1)
        panel.columnconfigure(0, weight=1)
        ttk.Label(
            panel,
            text='One name per line, e.g. Surname, Given.\nUsed for the Reviewer(s) filter and\nthe Actions "Assigned by" check.',
            style="Hint.TLabel",
            justify="left",
        ).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 6))
        self.reviewers = tk.Text(
            panel,
            width=28,
            height=10,
            wrap="none",
            font=("Segoe UI", 10),
            undo=True,
            relief="solid",
            borderwidth=1,
        )
        self.reviewers.grid(row=1, column=0, sticky="nsew")
        scroll_y = ttk.Scrollbar(panel, orient="vertical", command=self.reviewers.yview)
        scroll_y.grid(row=1, column=1, sticky="ns")
        scroll_x = ttk.Scrollbar(
            panel, orient="horizontal", command=self.reviewers.xview
        )
        scroll_x.grid(row=2, column=0, sticky="ew")
        self.reviewers.configure(
            yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set
        )
        self.reviewers.bind("<<Modified>>", self.names_modified)
        buttons = ttk.Frame(panel)
        buttons.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(6, 0))
        ttk.Button(buttons, text="Save list", command=self.save_reviewers).pack(
            side="left"
        )
        ttk.Button(buttons, text="Reload settings", command=self.reload_reviewers).pack(
            side="left", padx=(6, 0)
        )
        self.reviewer_status = tk.StringVar()
        ttk.Label(panel, textvariable=self.reviewer_status, style="Muted.TLabel").grid(
            row=4, column=0, columnspan=2, sticky="w", pady=(4, 0)
        )
        if "reviewer_names" in self.preferences:
            self.reload_reviewers()
        else:
            names = parse_names(
                "\n".join(
                    self.config["reviewer_names"] + self.config["assigned_by_names"]
                )
            )
            self.reviewers.insert("1.0", "\n".join(names))
            self.count_names()

    def build_documents_tab(self):
        self.load_table_filters()
        self.available_columns = {column[0] for column in DOCUMENT_COLUMNS}
        tab = ttk.Frame(self.tabs, padding=6)
        self.tabs.add(tab, text="Submissions")
        tab.columnconfigure(0, weight=1)
        tab.rowconfigure(1, weight=1)
        bar = ttk.Frame(tab)
        bar.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 6))
        bar.columnconfigure(3, weight=1)
        ttk.Label(bar, text="Show").grid(row=0, column=0, padx=(0, 4))
        self.result_filter = tk.StringVar(value="All")
        result_picker = ttk.Combobox(
            bar,
            textvariable=self.result_filter,
            values=RESULT_FILTERS,
            state="readonly",
            width=11,
        )
        result_picker.grid(row=0, column=1)
        result_picker.bind(
            "<<ComboboxSelected>>", lambda _event: self.render_documents()
        )
        ttk.Label(bar, text="Search").grid(row=0, column=2, padx=(10, 4))
        self.search = tk.StringVar()
        self.search.trace_add("write", lambda *_: self.render_documents())
        ttk.Entry(bar, textvariable=self.search, width=16).grid(
            row=0, column=3, sticky="ew"
        )
        self.open_button = ttk.Button(
            bar, text="Open link", command=self.open_link, state="disabled"
        )
        self.open_button.grid(row=0, column=4, padx=(8, 0))
        self.copy_button = ttk.Button(
            bar, text="Copy link", command=self.copy_link, state="disabled"
        )
        self.copy_button.grid(row=0, column=5, padx=(6, 0))
        ttk.Button(bar, text="Clear filters", command=self.clear_document_filters).grid(
            row=1, column=0, columnspan=2, sticky="w", pady=(4, 0)
        )
        self.documents = ttk.Treeview(
            tab,
            columns=[c[0] for c in DOCUMENT_COLUMNS],
            show="headings",
            selectmode="browse",
        )
        for index, (column, title, width) in enumerate(DOCUMENT_COLUMNS):
            self.documents.heading(
                column,
                text=f"{title} ▾",
                command=lambda index=index: self.open_column_filter(index),
            )
            self.documents.column(
                column,
                width=self.px(width),
                minwidth=self.px(60),
                stretch=False,
                anchor="center" if column in {"result", "submitted"} else "w",
            )
        for result, colour in (("yes", "#E7F4EA"), ("no", "#FBE9E7")):
            self.documents.tag_configure(result, background=colour)
        self.row_font = tkfont.nametofont("TkDefaultFont")
        self.heading_font = tkfont.Font(family="Segoe UI Semibold", size=9)
        self.pending_font = tkfont.Font(font=self.row_font)
        self.pending_font.configure(weight="bold")
        self.documents.tag_configure(
            "pending", foreground="#8B0000", font=self.pending_font
        )
        self.fit_columns()
        self.documents.grid(row=1, column=0, sticky="nsew")
        self.add_scrollbars(tab, self.documents)
        self.documents.bind("<Double-1>", lambda _event: self.open_link())
        self.documents.bind("<Return>", lambda _event: self.open_link())
        self.documents.bind(
            "<<TreeviewSelect>>", lambda _event: self.update_link_buttons()
        )
        self.documents.bind("<Control-c>", lambda _event: self.copy_link())
        self.empty_hint = ttk.Label(tab, text="", style="Muted.TLabel")
        self.empty_hint.grid(row=3, column=0, sticky="w", pady=(4, 0))

    def build_managers_tab(self):
        tab = ttk.Frame(self.tabs, padding=6)
        self.tabs.add(tab, text="By Author Manager")
        tab.columnconfigure(0, weight=1)
        tab.rowconfigure(0, weight=1)
        self.table = ttk.Treeview(
            tab,
            columns=("manager", "organization", "qualified", "yes", "no", "unresolved"),
            show="headings",
        )
        for column, title, width in (
            ("manager", "Author Manager", 250),
            ("organization", "Organization", 260),
            ("qualified", "Qualified", 100),
            ("yes", "Yes", 70),
            ("no", "No", 70),
            ("unresolved", "Unresolved", 110),
        ):
            self.table.heading(column, text=title)
            self.table.column(
                column,
                width=self.px(width),
                anchor="w" if column == "manager" else "center",
            )
        self.table.grid(row=0, column=0, sticky="nsew")
        scroll = ttk.Scrollbar(tab, orient="vertical", command=self.table.yview)
        scroll.grid(row=0, column=1, sticky="ns")
        self.table.configure(yscrollcommand=scroll.set)
        horizontal = ttk.Scrollbar(tab, orient="horizontal", command=self.table.xview)
        horizontal.grid(row=1, column=0, sticky="ew")
        self.table.configure(xscrollcommand=horizontal.set)
        self.table.bind("<Double-1>", lambda _event: self.show_manager())

    @staticmethod
    def add_scrollbars(parent, tree):
        scroll_y = ttk.Scrollbar(parent, orient="vertical", command=tree.yview)
        scroll_y.grid(row=1, column=1, sticky="ns")
        scroll_x = ttk.Scrollbar(parent, orient="horizontal", command=tree.xview)
        scroll_x.grid(row=2, column=0, sticky="ew")
        tree.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)

    def names_modified(self, _event):
        self.reviewers.edit_modified(False)
        self.count_names()

    def count_names(self):
        count = len(parse_names(self.reviewers.get("1.0", "end")))
        self.reviewer_status.set(f"{count} name{'s' if count != 1 else ''}")

    def column_value(self, row, index):
        value = (row[index] or "Unresolved") if index == 0 else row[index]
        if DOCUMENT_COLUMNS[index][0] == "submitted" and self.date_modes.get(index) == "month":
            return value[:7]
        return value

    def load_table_filters(self):
        self.column_filters, self.date_modes = {}, {}
        for index, (_, title, _) in enumerate(DOCUMENT_COLUMNS):
            selection = self.preferences.get("column_filters", {}).get(title)
            if selection is not None:
                if not selection.get("all", False):
                    self.column_filters[index] = set(selection["values"])
                if title == "Submitted":
                    self.date_modes[index] = "month" if selection.get("mode") == "month" else "day"

    def save_table_filters(self):
        filters = {
            DOCUMENT_COLUMNS[index][1]: {
                "mode": self.date_modes.get(index, "day") if DOCUMENT_COLUMNS[index][0] == "submitted" else "value",
                "values": sorted(values),
            }
            for index, values in self.column_filters.items()
        }
        for index, mode in self.date_modes.items():
            if index not in self.column_filters:
                filters[DOCUMENT_COLUMNS[index][1]] = {"mode": mode, "values": [], "all": True}
        try:
            save_preferences({"column_filters": filters})
        except (OSError, ValueError) as exc:
            messagebox.showerror("Cannot save filters", str(exc), parent=self.root)

    def open_column_filter(self, index):
        self.close_column_filter()
        if DOCUMENT_COLUMNS[index][0] not in self.available_columns:
            return None
        _, title, _ = DOCUMENT_COLUMNS[index]
        values = {self.column_value(row, index) for row in self.documents_data}
        values.update(self.column_filters.get(index, ()))
        popup = ttk.Frame(self.root, relief="solid", borderwidth=1)
        self.filter_popup = popup
        self.root.update_idletasks()
        widths = [
            int(self.documents.column(column, "width"))
            for column, _, _ in DOCUMENT_COLUMNS
            if column in self.available_columns
        ]
        visible_index = sum(column in self.available_columns for column, _, _ in DOCUMENT_COLUMNS[:index])
        offset = sum(widths[:visible_index]) - self.documents.xview()[0] * sum(widths)
        width = min(self.px(380), self.root.winfo_width() - 16)
        left = self.documents.winfo_rootx() - self.root.winfo_rootx() + round(offset)
        left = max(8, min(left, self.root.winfo_width() - width - 8))
        top = (
            self.documents.winfo_rooty()
            - self.root.winfo_rooty()
            + self.heading_font.metrics("linespace")
            + self.px(10)
        )
        height = min(self.px(440), self.root.winfo_height() - top - 8)
        popup.place(x=left, y=top, width=width, height=max(self.px(180), height))
        popup.lift()
        popup.columnconfigure(0, weight=1)
        popup.rowconfigure(3, weight=1)
        ttk.Label(popup, text=title, font=("Segoe UI Semibold", 10)).grid(
            row=0, column=0, sticky="w", padx=8, pady=(6, 0)
        )
        if DOCUMENT_COLUMNS[index][0] == "submitted":
            mode = tk.StringVar(value=self.date_modes.get(index, "day").title())
            picker = ttk.Combobox(popup, textvariable=mode, values=("Month", "Day"), state="readonly", width=8)
            picker.grid(row=0, column=1, padx=8, pady=(6, 0))

            def change_mode(_event):
                self.date_modes[index] = mode.get().lower()
                self.column_filters.pop(index, None)
                self.save_table_filters()
                self.open_column_filter(index)
                self.render_documents()

            picker.bind("<<ComboboxSelected>>", change_mode)
        query = tk.StringVar()
        entry = ttk.Entry(popup, textvariable=query)
        entry.grid(row=1, column=0, columnspan=2, sticky="ew", padx=8, pady=8)
        selected = self.column_filters.get(index, values)
        choices = {
            value: tk.BooleanVar(value=value in selected)
            for value in sorted(values, key=normalized)
        }
        buttons = ttk.Frame(popup)
        buttons.grid(row=2, column=0, columnspan=2, sticky="ew", padx=8)

        def select_visible(state):
            for value, variable in choices.items():
                if normalized(query.get()) in normalized(value or "(Blanks)"):
                    variable.set(state)

        ttk.Button(
            buttons, text="Select all", command=lambda: select_visible(True)
        ).pack(side="left")
        ttk.Button(
            buttons, text="Select none", command=lambda: select_visible(False)
        ).pack(side="left", padx=6)
        area = ttk.Frame(popup)
        area.grid(row=3, column=0, columnspan=2, sticky="nsew", padx=8, pady=8)
        area.columnconfigure(0, weight=1)
        area.rowconfigure(0, weight=1)
        canvas = tk.Canvas(area, highlightthickness=0)
        canvas.grid(row=0, column=0, sticky="nsew")
        scroll = ttk.Scrollbar(area, orient="vertical", command=canvas.yview)
        scroll.grid(row=0, column=1, sticky="ns")
        scroll_x = ttk.Scrollbar(area, orient="horizontal", command=canvas.xview)
        scroll_x.grid(row=1, column=0, sticky="ew")
        canvas.configure(yscrollcommand=scroll.set, xscrollcommand=scroll_x.set)
        options = ttk.Frame(canvas)
        window = canvas.create_window((0, 0), window=options, anchor="nw")
        canvas.bind(
            "<Configure>",
            lambda event: canvas.itemconfigure(
                window, width=max(event.width, options.winfo_reqwidth())
            ),
        )
        options.bind(
            "<Configure>",
            lambda _event: canvas.configure(scrollregion=canvas.bbox("all")),
        )
        popup.bind(
            "<MouseWheel>",
            lambda event: canvas.yview_scroll(-int(event.delta / 120), "units"),
        )
        checks = [
            (
                value,
                ttk.Checkbutton(options, text=value or "(Blanks)", variable=variable),
            )
            for value, variable in choices.items()
        ]

        def search_values(*_args):
            for value, check in checks:
                check.pack_forget()
                if normalized(query.get()) in normalized(value or "(Blanks)"):
                    check.pack(anchor="w", fill="x")
            canvas.yview_moveto(0)

        query.trace_add("write", search_values)
        search_values()

        def apply():
            chosen = {value for value, variable in choices.items() if variable.get()}
            if chosen == values:
                self.column_filters.pop(index, None)
            else:
                self.column_filters[index] = chosen
            self.save_table_filters()
            self.close_column_filter()
            self.render_documents()

        footer = ttk.Frame(popup)
        footer.grid(row=4, column=0, columnspan=2, sticky="ew", padx=8, pady=(0, 8))
        if DOCUMENT_COLUMNS[index][0] == "organization":

            def refresh_organizations():
                self.organization_refresh = True
                self.close_column_filter()
                self.start()

            ttk.Button(
                footer,
                text="Refresh organizations",
                command=refresh_organizations,
                state="disabled" if self.running else "normal",
            ).pack(side="left")
        ttk.Button(footer, text="Cancel", command=self.close_column_filter).pack(
            side="right"
        )
        ttk.Button(footer, text="Apply", command=apply).pack(side="right", padx=6)

        def outside_click(event):
            widget = event.widget
            while widget is not None:
                if widget == popup:
                    return
                widget = getattr(widget, "master", None)
            self.close_column_filter()

        def resize(event):
            if event.widget == self.root:
                self.close_column_filter()

        for sequence, callback in (
            ("<Button-1>", outside_click),
            ("<Escape>", lambda _event: self.close_column_filter()),
            ("<Configure>", resize),
            ("<Return>", lambda _event: apply()),
        ):
            self.filter_bindings.append(
                (sequence, self.root.bind(sequence, callback, add="+"))
            )
        entry.focus_set()
        return popup

    def close_column_filter(self):
        for sequence, binding in self.filter_bindings:
            self.root.unbind(sequence, binding)
        self.filter_bindings.clear()
        if self.filter_popup is not None:
            self.filter_popup.destroy()
            self.filter_popup = None

    def clear_document_filters(self):
        self.column_filters.clear()
        self.date_modes.clear()
        self.save_table_filters()
        self.result_filter.set("All")
        self.search.set("")
        self.render_documents()

    def show_document(self, row):
        result = row[0] or "Unresolved"
        wanted = self.result_filter.get()
        query = normalized(self.search.get())
        if (wanted != "All" and result != wanted) or (
            query and query not in normalized(" ".join(row))
        ):
            return
        if any(
            self.column_value(row, index) not in selected
            for index, selected in self.column_filters.items()
            if DOCUMENT_COLUMNS[index][0] in self.available_columns
        ):
            return
        self.documents.insert(
            "",
            "end",
            values=(result,) + row[1:],
            tags=(ROW_TAGS.get(row[0], "pending"),),
        )

    def fit_columns(self, rows=None, grow=False):
        pad = self.px(18)
        for index, (column, title, _) in enumerate(DOCUMENT_COLUMNS):
            heading = f"{title} {'●' if index in self.column_filters else '▾'}"
            self.documents.heading(column, text=heading)
            width = self.heading_font.measure(heading)
            for row in self.documents_data if rows is None else rows:
                value = row[index] or ("Unresolved" if index == 0 else "")
                font = self.row_font if row[0] in ROW_TAGS else self.pending_font
                width = max(width, font.measure(value))
            width += pad
            if grow:
                width = max(width, int(self.documents.column(column, "width")))
            self.documents.column(column, width=width)

    def update_document_hint(self):
        total = len(self.documents_data)
        self.empty_hint.configure(
            text=(
                f"Showing {len(self.documents.get_children())} of {total} submissions. "
                "Double-click a row to open its xCCB workspace."
                if total
                else ""
            )
        )

    def add_document(self, row):
        self.documents_data.append(row)
        self.show_document(row)
        self.fit_columns([row], grow=True)
        self.update_document_hint()

    def render_documents(self):
        self.documents.delete(*self.documents.get_children())
        for row in self.documents_data:
            self.show_document(row)
        self.fit_columns()
        self.update_document_hint()
        self.update_link_buttons()

    def selected_url(self):
        selection = self.documents.selection()
        return self.documents.set(selection[0], "url") if selection else ""

    def update_link_buttons(self):
        state = "normal" if self.selected_url() else "disabled"
        self.open_button.configure(state=state)
        self.copy_button.configure(state=state)

    def open_link(self):
        url = self.selected_url()
        try:
            workspace_identity(url)
        except ValueError:
            if url:
                messagebox.showwarning(
                    "Link not opened",
                    f"Not an xCCB Document Workspace link:\n{url}",
                    parent=self.root,
                )
            return
        webbrowser.open(url)

    def copy_link(self):
        url = self.selected_url()
        if url:
            self.root.clipboard_clear()
            self.root.clipboard_append(url)
            self.status.set("Link copied to clipboard")

    def show_manager(self):
        selection = self.table.selection()
        if selection:
            self.result_filter.set("All")
            self.search.set(self.table.set(selection[0], "manager"))
            self.tabs.select(0)

    def export(self):
        if not self.report_bytes:
            return
        path = filedialog.asksaveasfilename(
            parent=self.root,
            defaultextension=".xlsx",
            filetypes=[("Excel workbook", "*.xlsx")],
            initialfile=f"Flawless_Submission_{self.run_month}.xlsx",
        )
        if not path:
            return
        target = Path(path).resolve()
        if (
            self.source.get().strip()
            and target == Path(self.source.get().strip()).resolve()
        ):
            messagebox.showerror(
                "Export failed",
                "Choose a different file; the source workbook is never overwritten.",
                parent=self.root,
            )
            return
        temp = target.with_name(f".{target.name}.tmp")
        try:
            temp.write_bytes(self.report_bytes)
            os.replace(temp, target)
        except OSError as exc:
            temp.unlink(missing_ok=True)
            messagebox.showerror("Export failed", str(exc), parent=self.root)
            return
        self.status.set(f"Exported {target}")

    def choose_source(self):
        path = filedialog.askopenfilename(
            parent=self.root, filetypes=[("Excel workbook", "*.xlsx")]
        )
        if path:
            try:
                self.configure_workbook_columns(path)
            except (OSError, ValueError) as exc:
                messagebox.showerror("Upload failed", str(exc), parent=self.root)
                return
            self.source.set(path)
            self.status.set("Ready")

    def configure_workbook_columns(self, path):
        book, _, headers = read_source(path)
        book.close()
        sources = {
            "manager": "Author Manager", "phase": "Phase",
            "submitted": "Document Submitted", "url": "Document WS URL",
        }
        self.available_columns = {
            column for column, _, _ in DOCUMENT_COLUMNS
            if column not in sources or sources[column] in headers
        }
        self.documents.configure(displaycolumns=[
            column for column, _, _ in DOCUMENT_COLUMNS if column in self.available_columns
        ])
        self.render_documents()

    def save_reviewers(self):
        try:
            names = parse_names(self.reviewers.get("1.0", "end"))
            save_preferences({"reviewer_names": names})
        except (OSError, ValueError) as exc:
            messagebox.showerror("Cannot save listed names", str(exc), parent=self.root)
            return None
        self.reviewer_status.set(
            f"Saved {len(names)} name{'s' if len(names) != 1 else ''}"
        )
        return names

    def request_directory_runtime(self, reason):
        answer = Queue()
        self.events.put(("organization_setup", (reason, answer)))
        while not self.cancel.is_set():
            try:
                return answer.get(timeout=0.1)
            except Empty:
                pass
        return None

    def reload_reviewers(self):
        try:
            self.preferences = load_preferences()
            names = self.preferences.get("reviewer_names", [])
        except (OSError, ValueError) as exc:
            messagebox.showerror("Cannot read listed names", str(exc), parent=self.root)
            return
        self.reviewers.delete("1.0", "end")
        self.reviewers.insert("1.0", "\n".join(names))
        if hasattr(self, "documents"):
            self.load_table_filters()
            self.render_documents()
        self.reviewer_status.set(
            f"Loaded {len(names)} name{'s' if len(names) != 1 else ''}"
        )

    def start(self):
        if self.running:
            return
        self.close_column_filter()
        if self.organization_error:
            messagebox.showerror(
                "Organization error", self.organization_error, parent=self.root
            )
            return
        source, evidence = self.source.get().strip(), str(self.evidence)
        if not source or not Path(source).is_file():
            messagebox.showerror(
                "Choose workbook",
                "Choose an existing source workbook.",
                parent=self.root,
            )
            return
        selected_month = None
        try:
            self.configure_workbook_columns(source)
        except (OSError, ValueError) as exc:
            messagebox.showerror("Upload failed", str(exc), parent=self.root)
            return
        listed_names = self.save_reviewers()
        if listed_names is None:
            return
        config = self.config | {
            "reviewer_names": listed_names,
            "assigned_by_names": listed_names,
            "action_scope": "document",
        }
        self.run_month = "all_dates"
        self.running = True
        self.organization_warning = ""
        self.run_indicator.configure(text="Running", style="Running.TLabel")
        self.cancel = Event()
        self.report_bytes = None
        self.run_button.configure(state="disabled")
        self.stop_button.configure(state="normal")
        self.export_button.configure(state="disabled")
        self.table.delete(*self.table.get_children())
        self.manager_items.clear()
        self.documents_data = []
        self.render_documents()
        for value in self.card_values.values():
            value.set("…")
        self.status.set("Reading xCCB…")
        self.progress.grid()
        self.progress.start()
        Thread(
            target=self.generate,
            args=(source, evidence, config, selected_month, self.cancel),
            daemon=True,
        ).start()

    def stop(self):
        if self.running and not self.cancel.is_set():
            self.cancel.set()
            self.stop_button.configure(state="disabled")
            self.status.set("Stopping - waiting for in-flight xCCB requests…")

    def generate(self, source, evidence, config, selected_month, cancel):
        # Workers never touch Tk widgets; all UI changes run on the main thread.
        try:
            with TemporaryDirectory(prefix="xccb_dashboard_") as directory:
                output = Path(directory) / "report.xlsx"
                counts = report(
                    source,
                    output,
                    config,
                    evidence,
                    reader=WindowsReader(cancel=cancel),
                    on_progress=lambda counts, group, document: self.events.put(
                        ("progress", (counts, group, document))
                    ),
                    reporting_month=selected_month,
                    organization_codes=[],
                    refresh_organizations=self.organization_refresh,
                    on_organization_progress=lambda text: self.events.put(
                        ("organization_progress", text)
                    ),
                    on_directory_unavailable=(
                        None
                        if getattr(sys, "frozen", False)
                        else self.request_directory_runtime
                    ),
                    cancel=cancel,
                    all_dates=True,
                )
                status, groups, documents = read_results(output)
                data = output.read_bytes()
            self.events.put(("complete", (counts, status, groups, documents, data)))
        except Exception as exc:
            self.events.put(("error", str(exc)))

    def display_results(self, counts, groups):
        for _, key in CARDS:
            self.card_values[key].set(str(counts.get(key, 0)))
        for group in groups:
            manager = group[0]
            if manager in self.manager_items:
                self.table.item(self.manager_items[manager], values=group)
            else:
                self.manager_items[manager] = self.table.insert("", "end", values=group)
        for index, manager in enumerate(
            sorted(
                self.manager_items,
                key=lambda name: (name == "(Not specified)", normalized(name)),
            )
        ):
            self.table.move(self.manager_items[manager], "", index)

    def poll(self):
        for _ in range(100):
            try:
                event, result = self.events.get_nowait()
            except Empty:
                break
            if event == "organization_progress":
                self.status.set(result)
                if result.startswith("Organization lookup warning:"):
                    self.organization_warning = "Organization lookup issues; see Audit."
                continue
            if event == "organization_setup":
                reason, answer = result
                selected = ""
                if not self.cancel.is_set() and messagebox.askyesno(
                    "Organization lookup setup",
                    reason
                    + "\n\nChoose a Python installation with DataSyncX to retry?",
                    parent=self.root,
                ):
                    selected = filedialog.askopenfilename(
                        parent=self.root,
                        title="Select Python with DataSyncX",
                        filetypes=[("Python executable", "python.exe")],
                    )
                answer.put(selected or None)
                continue
            if event == "progress":
                counts, group, document = result
                self.display_results(counts, [group] if group is not None else [])
                if document is not None:
                    self.add_document(document_row(**document))
                if self.cancel.is_set():
                    continue
                self.status.set(
                    f"Processing - {sum(counts.get(key, 0) for key in ('Qualified', 'Excluded', 'Unverified'))} rows checked"
                )
                continue
            self.running = False
            self.run_indicator.configure(text="Stopped", style="Stopped.TLabel")
            self.run_button.configure(state="normal")
            self.stop_button.configure(state="disabled")
            self.progress.stop()
            self.progress.grid_remove()
            if event == "complete":
                self.organization_refresh = False
                counts, status, groups, documents, data = result
                self.status.set(
                    "DRAFT - unresolved results"
                    if status.startswith("DRAFT")
                    else status
                )
                if not counts.get("Qualified", 0):
                    self.status.set(
                        "No matching submissions. Check listed names and filters. "
                        f"Unverified eligibility: {counts.get('Unverified', 0)}."
                    )
                if self.organization_warning:
                    self.status.set(
                        self.status.get() + "; " + self.organization_warning
                    )
                self.display_results(counts, groups)
                self.documents_data = documents
                self.render_documents()
                self.report_bytes = data
                self.export_button.configure(state="normal")
                if self.cancel.is_set():
                    stopped = sum(row[-1] == "Stopped" for row in documents)
                    self.status.set(
                        f"Stopped by user - {stopped} submission{'s' if stopped != 1 else ''} not checked."
                    )
            else:
                self.status.set("Report failed. " + result)
                messagebox.showerror("Report failed", result, parent=self.root)
        self.root.after(100, self.poll)

    def close(self):
        if self.running:
            messagebox.showinfo(
                "Report running",
                "Wait for the report to finish before closing.",
                parent=self.root,
            )
        else:
            self.root.destroy()


def main():
    if sys.platform == "win32":
        import ctypes

        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except (AttributeError, OSError):
            pass
    root = tk.Tk()
    style = ttk.Style(root)
    if "vista" in style.theme_names():
        style.theme_use("vista")
    Dashboard(root)
    root.mainloop()


if __name__ == "__main__":
    main()

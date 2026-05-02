from __future__ import annotations

import tkinter as tk
from dataclasses import dataclass
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Callable, Iterable


APP_TITLE = "EasyEdits"
ACTION_BUTTON_WIDTH = 24


@dataclass(frozen=True)
class RenameItem:
    source: Path
    target: Path

    @property
    def will_change(self) -> bool:
        return self.source != self.target


def collect_files(folder: Path, recursive: bool) -> list[Path]:
    iterator: Iterable[Path] = folder.rglob("*") if recursive else folder.glob("*")
    return sorted(path for path in iterator if path.is_file())


def validate_plan(plan: list[RenameItem]) -> None:
    changed = [item for item in plan if item.will_change]
    targets = [item.target.resolve() for item in changed]

    if len(set(targets)) != len(targets):
        raise ValueError("Two or more files would end up with the same name.")

    sources = {item.source.resolve() for item in changed}
    conflicts = [
        item.target.name
        for item in changed
        if item.target.exists() and item.target.resolve() not in sources
    ]
    if conflicts:
        names = "\n".join(f"  - {name}" for name in conflicts)
        raise FileExistsError(
            "Rename stopped because these target files already exist:\n" + names
        )


def apply_plan(plan: list[RenameItem]) -> list[RenameItem]:
    validate_plan(plan)
    changed = [item for item in plan if item.will_change]
    if not changed:
        return []

    temp_pairs: list[tuple[Path, Path]] = []
    final_pairs: list[tuple[Path, Path]] = []
    applied_final: list[tuple[Path, Path]] = []

    try:
        for index, item in enumerate(changed):
            temp = item.source.with_name(f".easyedits_tmp_{index}_{item.source.name}")
            while temp.exists():
                temp = temp.with_name(f".easyedits_tmp_{index}_{temp.name}")
            item.source.rename(temp)
            temp_pairs.append((temp, item.source))
            final_pairs.append((temp, item.target))

        for temp, target in final_pairs:
            temp.rename(target)
            applied_final.append((target, temp))
    except Exception:
        for target, temp in reversed(applied_final):
            if target.exists():
                target.rename(temp)
        for temp, source in reversed(temp_pairs):
            if temp.exists():
                temp.rename(source)
        raise

    return [RenameItem(source=item.target, target=item.source) for item in changed]


def add_example_lines(
    parent: ttk.Frame,
    start_row: int,
    lines: tuple[str, str, str],
    columnspan: int,
) -> None:
    for offset, line in enumerate(lines):
        ttk.Label(parent, text=line, foreground="dim gray").grid(
            row=start_row + offset,
            column=0,
            columnspan=columnspan,
            sticky="w",
            pady=((12 if offset == 0 else 2), 0),
        )


class EasyEditsApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title(APP_TITLE)

        self.folder_var = tk.StringVar(value="")
        self.recursive_var = tk.BooleanVar(value=False)
        self.status_var = tk.StringVar(value="Ready.")

        self.replace_find_var = tk.StringVar()
        self.replace_with_var = tk.StringVar()
        self.remove_text_var = tk.StringVar()
        self.add_text_var = tk.StringVar()
        self.add_position_var = tk.StringVar(value="Beginning")
        self.add_reference_var = tk.StringVar()
        self.sequence_name_var = tk.StringVar()
        self.sequence_start_var = tk.StringVar(value="1")
        self.sequence_padding_var = tk.StringVar(value="No padding")
        self.sequence_separator_var = tk.StringVar(value="Space")
        self.sequence_position_var = tk.StringVar(value="After name")
        self.original_name_var = tk.StringVar()
        self.single_name_var = tk.StringVar()

        self.files: list[Path] = []
        self.checked_sources: set[Path] = set()
        self.plan: list[RenameItem] = []
        self.undo_plan: list[RenameItem] = []
        self.manual_undo_plan: list[RenameItem] = []
        self.selected_source: Path | None = None

        self._build_ui()
        self.single_name_var.trace_add("write", self.on_single_name_change)
        self.autosize_window()

    def autosize_window(self) -> None:
        self.update_idletasks()
        width = min(max(self.winfo_reqwidth() + 20, 1100), self.winfo_screenwidth() - 80)
        height = min(max(self.winfo_reqheight() + 20, 720), self.winfo_screenheight() - 100)
        left = max((self.winfo_screenwidth() - width) // 2, 0)
        top = max((self.winfo_screenheight() - height) // 2, 0)
        self.geometry(f"{width}x{height}+{left}+{top}")
        self.minsize(900, 500)

    def _build_ui(self) -> None:
        shell = ttk.Frame(self)
        shell.pack(fill=tk.BOTH, expand=True)
        shell.columnconfigure(0, weight=1)
        shell.rowconfigure(0, weight=1)

        self.page_canvas = tk.Canvas(shell, highlightthickness=0)
        self.page_canvas.grid(row=0, column=0, sticky="nsew")
        page_scrollbar = ttk.Scrollbar(shell, orient=tk.VERTICAL, command=self.page_canvas.yview)
        page_scrollbar.grid(row=0, column=1, sticky="ns")
        self.page_canvas.configure(yscrollcommand=page_scrollbar.set)

        root = ttk.Frame(self.page_canvas, padding=14)
        root_window = self.page_canvas.create_window((0, 0), window=root, anchor="nw")

        def update_page_scroll_region(_event: tk.Event | None = None) -> None:
            self.page_canvas.configure(scrollregion=self.page_canvas.bbox("all"))

        def update_page_width(event: tk.Event) -> None:
            self.page_canvas.itemconfigure(root_window, width=event.width)

        root.bind("<Configure>", update_page_scroll_region)
        self.page_canvas.bind("<Configure>", update_page_width)
        self.page_canvas.bind_all("<MouseWheel>", self.on_mousewheel)

        root.columnconfigure(0, weight=1)
        root.rowconfigure(4, weight=1)

        header = ttk.Frame(root)
        header.grid(row=0, column=0, sticky="ew")
        header.columnconfigure(0, weight=1)
        ttk.Label(header, text="EasyEdits", font=("Segoe UI", 18, "bold")).grid(
            row=0, column=0, sticky="w"
        )

        folder_row = ttk.Frame(root)
        folder_row.grid(row=1, column=0, sticky="ew", pady=(12, 10))
        folder_row.columnconfigure(1, weight=1)
        ttk.Label(folder_row, text="Directory").grid(row=0, column=0, padx=(0, 8))
        ttk.Entry(folder_row, textvariable=self.folder_var).grid(row=0, column=1, sticky="ew")
        ttk.Button(folder_row, text="Browse", command=self.browse_folder).grid(
            row=0, column=2, padx=(8, 0)
        )
        ttk.Button(folder_row, text="Refresh", command=self.refresh_files).grid(
            row=0, column=3, padx=(8, 0)
        )
        ttk.Checkbutton(
            folder_row, text="Include subfolders", variable=self.recursive_var, command=self.refresh_files
        ).grid(row=0, column=4, padx=(12, 0))

        self.tabs = ttk.Notebook(root)
        self.tabs.grid(row=2, column=0, sticky="ew", pady=(0, 10))
        self._build_replace_tab()
        self._build_remove_tab()
        self._build_add_tab()
        self._build_sequence_tab()

        list_toolbar = ttk.Frame(root)
        list_toolbar.grid(row=3, column=0, sticky="ew", pady=(0, 6))
        list_toolbar.columnconfigure(3, weight=1)
        ttk.Label(list_toolbar, text="Files And Folders", font=("Segoe UI", 10, "bold")).grid(
            row=0, column=0, sticky="w"
        )
        ttk.Button(list_toolbar, text="Check All", command=self.check_all).grid(
            row=0, column=1, padx=(16, 8)
        )
        ttk.Button(list_toolbar, text="Clear Checks", command=self.clear_checks).grid(
            row=0, column=2, sticky="w"
        )

        content = ttk.Frame(root)
        content.grid(row=4, column=0, sticky="nsew")
        content.columnconfigure(0, weight=1)
        content.rowconfigure(0, weight=1)

        grid_frame = ttk.Frame(content)
        grid_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 14))
        grid_frame.columnconfigure(0, weight=1)
        grid_frame.rowconfigure(0, weight=1)

        columns = ("use", "current", "new", "status")
        self.tree = ttk.Treeview(grid_frame, columns=columns, show="headings", selectmode="browse")
        self.tree.heading("use", text="Use")
        self.tree.heading("current", text="Current Name")
        self.tree.heading("new", text="New Name")
        self.tree.heading("status", text="Status")
        self.tree.column("use", width=60, anchor=tk.CENTER, stretch=False)
        self.tree.column("current", width=400, anchor=tk.W)
        self.tree.column("new", width=400, anchor=tk.W)
        self.tree.column("status", width=120, anchor=tk.W, stretch=False)
        self.tree.grid(row=0, column=0, sticky="nsew")
        self.tree.bind("<Double-1>", self.toggle_selected_row)
        self.tree.bind("<space>", self.toggle_selected_row)
        self.tree.bind("<<TreeviewSelect>>", self.on_tree_select)

        scrollbar = ttk.Scrollbar(grid_frame, orient=tk.VERTICAL, command=self.tree.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=scrollbar.set)

        details = ttk.Frame(content)
        details.grid(row=0, column=1, sticky="nsew")
        details.columnconfigure(0, weight=1)
        details.rowconfigure(4, weight=1)
        ttk.Label(details, text="Selected File", font=("Segoe UI", 10, "bold")).grid(
            row=0, column=0, sticky="w"
        )
        ttk.Label(details, text="Original Name").grid(row=1, column=0, sticky="w", pady=(12, 3))
        ttk.Entry(details, textvariable=self.original_name_var, width=42, state="readonly").grid(
            row=2, column=0, sticky="ew"
        )
        ttk.Label(details, text="Changed Text").grid(row=3, column=0, sticky="w", pady=(12, 3))
        changed_text_frame = ttk.Frame(details)
        changed_text_frame.grid(row=4, column=0, sticky="nsew")
        changed_text_frame.columnconfigure(0, weight=1)
        changed_text_frame.rowconfigure(0, weight=1)
        self.changed_text = tk.Text(
            changed_text_frame,
            width=48,
            height=13,
            wrap="none",
            relief=tk.SOLID,
            borderwidth=1,
        )
        self.changed_text.grid(row=0, column=0, sticky="nsew")
        changed_y = ttk.Scrollbar(changed_text_frame, orient=tk.VERTICAL, command=self.changed_text.yview)
        changed_y.grid(row=0, column=1, sticky="ns")
        changed_x = ttk.Scrollbar(changed_text_frame, orient=tk.HORIZONTAL, command=self.changed_text.xview)
        changed_x.grid(row=1, column=0, sticky="ew")
        self.changed_text.configure(yscrollcommand=changed_y.set, xscrollcommand=changed_x.set)
        self.changed_text.configure(state="disabled")
        ttk.Label(details, text="Single File Edit", font=("Segoe UI", 10, "bold")).grid(
            row=5, column=0, sticky="w", pady=(14, 3)
        )
        ttk.Label(details, text="New Name").grid(row=6, column=0, sticky="w", pady=(4, 3))
        ttk.Entry(details, textvariable=self.single_name_var, width=42).grid(row=7, column=0, sticky="ew")
        manual_buttons = ttk.Frame(details)
        manual_buttons.grid(row=8, column=0, sticky="ew", pady=(8, 0))
        manual_buttons.columnconfigure(0, weight=1)
        ttk.Button(
            manual_buttons,
            text="Apply Single File Rename",
            command=self.apply_single_rename,
            width=ACTION_BUTTON_WIDTH,
        ).grid(row=0, column=0, sticky="ew", pady=(0, 6))
        ttk.Button(
            manual_buttons,
            text="Undo Single File Rename",
            command=self.undo_single_rename,
            width=ACTION_BUTTON_WIDTH,
        ).grid(row=1, column=0, sticky="ew")

        footer = ttk.Frame(root)
        footer.grid(row=5, column=0, sticky="ew", pady=(10, 0))
        footer.columnconfigure(0, weight=1)
        ttk.Label(footer, textvariable=self.status_var, anchor=tk.W).grid(
            row=0, column=0, sticky="ew"
        )

    def on_mousewheel(self, event: tk.Event) -> None:
        self.page_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def add_action_buttons(
        self,
        tab: ttk.Frame,
        column: int,
        preview_command: Callable[[], None],
    ) -> None:
        ttk.Button(
            tab,
            text="Preview Changes",
            command=preview_command,
            width=ACTION_BUTTON_WIDTH,
        ).grid(row=0, column=column, rowspan=1, sticky="e")
        ttk.Button(
            tab,
            text="Apply Rename",
            command=self.apply_changes,
            width=ACTION_BUTTON_WIDTH,
        ).grid(
            row=1, column=column, sticky="e", pady=(6, 0)
        )
        ttk.Button(
            tab,
            text="Undo Changes",
            command=self.undo_changes,
            width=ACTION_BUTTON_WIDTH,
        ).grid(
            row=2, column=column, sticky="e", pady=(6, 0)
        )

    def _build_replace_tab(self) -> None:
        tab = ttk.Frame(self.tabs, padding=12)
        self.tabs.add(tab, text="Replace")
        tab.columnconfigure(1, weight=1)
        tab.columnconfigure(3, weight=1)
        ttk.Label(tab, text="Filename Text To Be Replaced").grid(row=0, column=0, sticky="w")
        ttk.Entry(tab, textvariable=self.replace_find_var).grid(
            row=0, column=1, sticky="ew", padx=(8, 18)
        )
        ttk.Label(tab, text="Filename Text Replaced With").grid(row=0, column=2, sticky="w")
        ttk.Entry(tab, textvariable=self.replace_with_var).grid(
            row=0, column=3, sticky="ew", padx=(8, 18)
        )
        self.add_action_buttons(tab, 4, self.preview_replace)
        add_example_lines(
            tab,
            1,
            (
                "Example: Original File Name - report_draft.txt",
                "Example: Filename Text To Be Replaced = 'draft' -> Filename Text Replaced With = 'final'",
                "Example: End Result - report_draft.txt -> report_final.txt",
            ),
            4,
        )

    def _build_remove_tab(self) -> None:
        tab = ttk.Frame(self.tabs, padding=12)
        self.tabs.add(tab, text="Remove")
        tab.columnconfigure(1, weight=1)
        ttk.Label(tab, text="Remove Text").grid(row=0, column=0, sticky="w")
        ttk.Entry(tab, textvariable=self.remove_text_var).grid(
            row=0, column=1, sticky="ew", padx=(8, 18)
        )
        self.add_action_buttons(tab, 2, self.preview_remove)
        add_example_lines(
            tab,
            1,
            (
                "Example: Original File Name - summer_EDIT_photo.jpg",
                "Example: Remove Text = 'EDIT_'",
                "Example: End Result - summer_EDIT_photo.jpg -> summer_photo.jpg",
            ),
            2,
        )

    def _build_add_tab(self) -> None:
        tab = ttk.Frame(self.tabs, padding=12)
        self.tabs.add(tab, text="Add")
        tab.columnconfigure(1, weight=1)
        ttk.Label(tab, text="Text To Add").grid(row=0, column=0, sticky="w")
        ttk.Entry(tab, textvariable=self.add_text_var).grid(
            row=0, column=1, sticky="ew", padx=(8, 18)
        )
        ttk.Label(tab, text="Add Position").grid(row=0, column=2, sticky="w")
        ttk.Combobox(
            tab,
            textvariable=self.add_position_var,
            values=("Beginning", "End", "Before reference text", "After reference text"),
            state="readonly",
            width=22,
        ).grid(row=0, column=3, sticky="w", padx=(8, 18))
        ttk.Label(tab, text="Reference Text").grid(row=0, column=4, sticky="w")
        ttk.Entry(tab, textvariable=self.add_reference_var, width=14).grid(
            row=0, column=5, sticky="w", padx=(8, 18)
        )
        self.add_action_buttons(tab, 6, self.preview_add)
        add_example_lines(
            tab,
            1,
            (
                "Example: Original File Name - report_final.txt",
                "Example: Text To Add = 'v2_', Add Position = 'Beginning'",
                "Example: End Result - report_final.txt -> v2_report_final.txt",
            ),
            6,
        )

    def _build_sequence_tab(self) -> None:
        outer = ttk.Frame(self.tabs)
        self.tabs.add(outer, text="Sequential Rename")
        outer.columnconfigure(0, weight=1)

        canvas = tk.Canvas(outer, height=160, highlightthickness=0)
        canvas.grid(row=0, column=0, sticky="ew")
        scrollbar = ttk.Scrollbar(outer, orient=tk.HORIZONTAL, command=canvas.xview)
        scrollbar.grid(row=1, column=0, sticky="ew")
        canvas.configure(xscrollcommand=scrollbar.set)

        tab = ttk.Frame(canvas, padding=12)
        canvas_window = canvas.create_window((0, 0), window=tab, anchor="nw")

        def update_scroll_region(_event: tk.Event | None = None) -> None:
            canvas.configure(scrollregion=canvas.bbox("all"))

        def update_canvas_width(event: tk.Event) -> None:
            requested_width = tab.winfo_reqwidth()
            canvas.itemconfigure(canvas_window, width=max(requested_width, event.width))

        tab.bind("<Configure>", update_scroll_region)
        canvas.bind("<Configure>", update_canvas_width)

        ttk.Label(tab, text="Rename Checked In Order As").grid(row=0, column=0, sticky="w")
        ttk.Entry(tab, textvariable=self.sequence_name_var, width=28).grid(
            row=0, column=1, sticky="w", padx=(8, 18)
        )
        self.add_action_buttons(tab, 8, self.preview_sequence)
        ttk.Label(tab, text="Start Number").grid(row=1, column=0, sticky="w", pady=(12, 0))
        ttk.Entry(tab, textvariable=self.sequence_start_var, width=8).grid(
            row=1, column=1, sticky="w", padx=(8, 18), pady=(12, 0)
        )
        ttk.Label(tab, text="Padding").grid(row=1, column=2, sticky="w", pady=(12, 0))
        ttk.Combobox(
            tab,
            textvariable=self.sequence_padding_var,
            values=("No padding", "2 digits", "3 digits", "4 digits"),
            state="readonly",
            width=12,
        ).grid(row=1, column=3, sticky="w", padx=(8, 18), pady=(12, 0))
        ttk.Label(tab, text="Separator").grid(row=1, column=4, sticky="w", pady=(12, 0))
        ttk.Combobox(
            tab,
            textvariable=self.sequence_separator_var,
            values=("Space", "Dash (-)", "Dash ( - )", "Underscore (_)", "None"),
            state="readonly",
            width=15,
        ).grid(row=1, column=5, sticky="w", padx=(8, 18), pady=(12, 0))
        ttk.Label(tab, text="Number Position").grid(row=1, column=6, sticky="w", pady=(12, 0))
        ttk.Combobox(
            tab,
            textvariable=self.sequence_position_var,
            values=("After name", "After name (no space)", "Before name"),
            state="readonly",
            width=20,
        ).grid(row=1, column=7, sticky="w", padx=(8, 18), pady=(12, 0))
        add_example_lines(
            tab,
            2,
            (
                "Example: Original File Name - IMG_4321.jpg, IMG_4322.jpg, IMG_4323.jpg",
                "Example: Rename Checked In Order As = 'Vacation', Start Number = '1', Padding = '2 digits', Separator = 'Space', Number Position = 'After name'",
                "Example: End Result - Vacation 01.jpg, Vacation 02.jpg, Vacation 03.jpg",
            ),
            8,
        )

    def browse_folder(self) -> None:
        folder = filedialog.askdirectory(
            initialdir=self.folder_var.get() or str(Path.cwd()),
            mustexist=True,
        )
        if folder:
            self.folder_var.set(folder)
            self.refresh_files()

    def refresh_files(self) -> None:
        try:
            folder_text = self.folder_var.get().strip()
            if not folder_text:
                self.status_var.set("Choose a folder with Browse first.")
                return

            folder = Path(folder_text).expanduser().resolve()
            if not folder.exists() or not folder.is_dir():
                raise FileNotFoundError(f"Folder not found: {folder}")

            self.files = collect_files(folder, self.recursive_var.get())
            self.checked_sources = set()
            self.plan = [RenameItem(path, path) for path in self.files]
            self.show_plan()
            self.status_var.set(f"Loaded {len(self.files)} file(s).")
        except Exception as exc:
            self.files = []
            self.checked_sources = set()
            self.plan = []
            self.show_plan()
            self.status_var.set(str(exc))
            messagebox.showerror(APP_TITLE, str(exc))

    def checked_files(self) -> list[Path]:
        return [path for path in self.files if path in self.checked_sources]

    def has_selected_folder(self) -> bool:
        if self.folder_var.get().strip():
            return True
        self.status_var.set("Choose a folder with Browse first.")
        return False

    def check_all(self) -> None:
        self.checked_sources = set(self.files)
        self.show_plan()
        self.update_status()

    def clear_checks(self) -> None:
        self.checked_sources = set()
        self.show_plan()
        self.update_status()

    def toggle_selected_row(self, event: tk.Event | None = None) -> str | None:
        selected = self.tree.selection()
        if not selected:
            return None
        source = Path(selected[0])
        if source in self.checked_sources:
            self.checked_sources.remove(source)
        else:
            self.checked_sources.add(source)
        self.show_plan()
        self.tree.selection_set(str(source))
        self.update_status()
        return "break"

    def on_tree_select(self, event: tk.Event | None = None) -> None:
        selected = self.tree.selection()
        self.selected_source = Path(selected[0]) if selected else None
        self.show_selection_details()

    def selected_plan_item(self) -> RenameItem | None:
        if self.selected_source is None:
            return None
        for item in self.plan:
            if item.source == self.selected_source:
                return item
        return None

    def show_selection_details(self) -> None:
        item = self.selected_plan_item()
        if item is None:
            self.original_name_var.set("")
            self.single_name_var.set("")
            self.set_changed_text("", "")
            return

        self.original_name_var.set(item.source.name)
        self.single_name_var.set(item.target.name)
        removed, added = change_parts(item.source.name, item.target.name)
        self.set_changed_text(removed, added)

    def on_single_name_change(self, *_args: object) -> None:
        if self.selected_source is None:
            return

        new_name = self.single_name_var.get()
        if not new_name:
            self.set_changed_text("", "")
            return

        for index, item in enumerate(self.plan):
            if item.source != self.selected_source:
                continue

            updated = RenameItem(item.source, item.source.with_name(new_name))
            self.plan[index] = updated
            removed, added = change_parts(updated.source.name, updated.target.name)
            self.set_changed_text(removed, added)
            if str(updated.source) in self.tree.get_children():
                checked = "[x]" if updated.source in self.checked_sources else "[ ]"
                status = "Will rename" if updated.will_change else "No change"
                self.tree.item(
                    str(updated.source),
                    values=(checked, updated.source.name, updated.target.name, status),
                )
            self.update_status()
            return

    def set_changed_text(self, removed: str, added: str) -> None:
        self.changed_text.configure(state="normal")
        self.changed_text.delete("1.0", tk.END)
        self.changed_text.insert(tk.END, "Removed:\n")
        self.changed_text.insert(tk.END, visible_text(removed) if removed else "(nothing removed)")
        self.changed_text.insert(tk.END, "\n\nAdded:\n")
        self.changed_text.insert(tk.END, visible_text(added) if added else "(nothing added)")
        self.changed_text.configure(state="disabled")

    def apply_single_rename(self) -> None:
        item = self.selected_plan_item()
        if item is None:
            self.status_var.set("Select a file first.")
            return

        new_name = self.single_name_var.get().strip()
        if not new_name:
            self.status_var.set("Type a new file name first.")
            return

        target = item.source.with_name(new_name)
        manual_plan = [RenameItem(item.source, target)]
        try:
            validate_plan(manual_plan)
            self.manual_undo_plan = apply_plan(manual_plan)
            self.refresh_files()
            self.status_var.set("Manually renamed 1 file.")
        except Exception as exc:
            self.status_var.set(str(exc))
            messagebox.showerror(APP_TITLE, str(exc))

    def undo_single_rename(self) -> None:
        if not self.manual_undo_plan:
            self.status_var.set("Nothing to undo for manual rename.")
            return

        try:
            apply_plan(self.manual_undo_plan)
            self.manual_undo_plan = []
            self.refresh_files()
            self.status_var.set("Undid single file rename.")
        except Exception as exc:
            self.status_var.set(str(exc))
            messagebox.showerror(APP_TITLE, str(exc))

    def update_status(self) -> None:
        checked = len(self.checked_sources)
        changed = sum(1 for item in self.plan if item.will_change)
        self.status_var.set(f"{checked} checked. {changed} staged rename(s).")

    def preview_replace(self) -> None:
        if not self.has_selected_folder():
            return
        find_text = self.replace_find_var.get()
        if not find_text:
            self.status_var.set("Type text into Replace Text first.")
            return
        replace_with = self.replace_with_var.get()
        self.preview_text_transform(
            lambda name: name.replace(find_text, replace_with),
            "Replaced text in checked files.",
        )

    def preview_remove(self) -> None:
        if not self.has_selected_folder():
            return
        remove_text = self.remove_text_var.get()
        if not remove_text:
            self.status_var.set("Type text into Remove Text first.")
            return
        self.preview_text_transform(
            lambda name: name.replace(remove_text, ""),
            "Removed text from checked files.",
        )

    def preview_add(self) -> None:
        if not self.has_selected_folder():
            return
        text_to_add = self.add_text_var.get()
        if not text_to_add:
            self.status_var.set("Type text into Text To Add first.")
            return

        position = self.add_position_var.get()
        reference = self.add_reference_var.get()

        try:
            self.preview_text_transform(
                lambda name: insert_text(name, text_to_add, position, reference),
                "Added text to checked files.",
            )
        except Exception as exc:
            self.status_var.set(str(exc))
            messagebox.showerror(APP_TITLE, str(exc))

    def preview_text_transform(
        self, transform: Callable[[str], str], success_message: str
    ) -> None:
        if not self.has_selected_folder():
            return
        try:
            self.plan = []
            for path in self.files:
                target_name = transform(path.name) if path in self.checked_sources else path.name
                self.plan.append(RenameItem(path, path.with_name(target_name)))
            validate_plan(self.plan)
            self.show_plan(prefer_changed=True)
            self.status_var.set(success_message)
        except Exception as exc:
            self.status_var.set(str(exc))
            messagebox.showerror(APP_TITLE, str(exc))

    def preview_sequence(self) -> None:
        if not self.has_selected_folder():
            return
        base_name = self.sequence_name_var.get()
        if not base_name:
            self.status_var.set("Type a name into Rename Checked In Order As first.")
            return

        try:
            start = int(self.sequence_start_var.get())
            padding = padding_length(self.sequence_padding_var.get())
            separator = separator_text(self.sequence_separator_var.get())
            position = self.sequence_position_var.get()

            selected = self.checked_files()
            targets: dict[Path, str] = {}
            for index, path in enumerate(selected):
                number = str(start + index).zfill(padding) if padding else str(start + index)
                if position == "Before name":
                    stem = f"{number}{separator}{base_name}"
                elif position == "After name (no space)":
                    stem = f"{base_name}{number}"
                else:
                    stem = f"{base_name}{separator}{number}"
                targets[path] = f"{stem}{path.suffix}"

            self.plan = [
                RenameItem(path, path.with_name(targets.get(path, path.name)))
                for path in self.files
            ]
            validate_plan(self.plan)
            self.show_plan(prefer_changed=True)
            self.status_var.set("Prepared sequential names for checked files.")
        except Exception as exc:
            self.status_var.set(str(exc))
            messagebox.showerror(APP_TITLE, str(exc))

    def apply_changes(self) -> None:
        changed_count = sum(1 for item in self.plan if item.will_change)
        if changed_count == 0:
            self.preview_active_tab()
            changed_count = sum(1 for item in self.plan if item.will_change)
            if changed_count == 0:
                self.status_var.set("Nothing to rename.")
                return
        if not messagebox.askyesno(APP_TITLE, f"Rename {changed_count} file(s)?"):
            return

        try:
            self.undo_plan = apply_plan(self.plan)
            self.refresh_files()
            self.status_var.set(f"Renamed {changed_count} file(s). Undo is available.")
        except Exception as exc:
            self.status_var.set(str(exc))
            messagebox.showerror(APP_TITLE, str(exc))

    def undo_changes(self) -> None:
        if not self.undo_plan:
            self.status_var.set("Nothing to undo.")
            return
        if not messagebox.askyesno(APP_TITLE, "Undo the last rename?"):
            return

        try:
            undo_count = len(self.undo_plan)
            apply_plan(self.undo_plan)
            self.undo_plan = []
            self.refresh_files()
            self.status_var.set(f"Undid {undo_count} rename(s).")
        except Exception as exc:
            self.status_var.set(str(exc))
            messagebox.showerror(APP_TITLE, str(exc))

    def preview_active_tab(self) -> None:
        selected_tab = self.tabs.tab(self.tabs.select(), "text")
        if selected_tab == "Replace":
            self.preview_replace()
        elif selected_tab == "Remove":
            self.preview_remove()
        elif selected_tab == "Add":
            self.preview_add()
        elif selected_tab == "Sequential Rename":
            self.preview_sequence()

    def show_plan(self, prefer_changed: bool = False) -> None:
        selected = self.selected_source
        for item_id in self.tree.get_children():
            self.tree.delete(item_id)

        for item in self.plan:
            checked = "[x]" if item.source in self.checked_sources else "[ ]"
            status = "Will rename" if item.will_change else "No change"
            self.tree.insert(
                "",
                tk.END,
                iid=str(item.source),
                values=(checked, item.source.name, item.target.name, status),
            )

        first_changed = next((item.source for item in self.plan if item.will_change), None)
        if prefer_changed and first_changed:
            self.selected_source = first_changed
            self.tree.selection_set(str(first_changed))
            self.tree.see(str(first_changed))
        elif selected and str(selected) in self.tree.get_children():
            self.tree.selection_set(str(selected))
        elif self.plan:
            self.selected_source = self.plan[0].source
            self.tree.selection_set(str(self.plan[0].source))
        else:
            self.selected_source = None
        self.show_selection_details()


def insert_text(name: str, text_to_add: str, position: str, reference: str) -> str:
    if position == "Beginning":
        return f"{text_to_add}{name}"
    if position == "End":
        return f"{name}{text_to_add}"
    if not reference:
        raise ValueError("Type reference text first.")

    index = name.find(reference)
    if index == -1:
        return name
    if position == "Before reference text":
        return name[:index] + text_to_add + name[index:]
    if position == "After reference text":
        insert_at = index + len(reference)
        return name[:insert_at] + text_to_add + name[insert_at:]
    return f"{name}{text_to_add}"


def change_parts(source_name: str, target_name: str) -> tuple[str, str]:
    prefix_length = 0
    max_prefix = min(len(source_name), len(target_name))
    while (
        prefix_length < max_prefix
        and source_name[prefix_length] == target_name[prefix_length]
    ):
        prefix_length += 1

    suffix_length = 0
    max_suffix = min(len(source_name), len(target_name)) - prefix_length
    while (
        suffix_length < max_suffix
        and source_name[len(source_name) - 1 - suffix_length]
        == target_name[len(target_name) - 1 - suffix_length]
    ):
        suffix_length += 1

    source_end = len(source_name) - suffix_length if suffix_length else len(source_name)
    target_end = len(target_name) - suffix_length if suffix_length else len(target_name)

    while (
        source_end < len(source_name)
        and target_end < len(target_name)
        and source_name[source_end] == target_name[target_end]
        and is_name_word_char(source_name[source_end])
    ):
        source_end += 1
        target_end += 1

    return source_name[prefix_length:source_end], target_name[prefix_length:target_end]


def is_name_word_char(value: str) -> bool:
    return value.isalnum()


def visible_text(value: str) -> str:
    return value.replace(" ", "[space]").replace("\t", "[tab]")


def padding_length(choice: str) -> int:
    return {
        "No padding": 0,
        "2 digits": 2,
        "3 digits": 3,
        "4 digits": 4,
    }.get(choice, 0)


def separator_text(choice: str) -> str:
    return {
        "Space": " ",
        "Dash (-)": "-",
        "Dash ( - )": " - ",
        "Underscore (_)": "_",
        "None": "",
    }.get(choice, " ")


def main() -> None:
    app = EasyEditsApp()
    app.mainloop()


if __name__ == "__main__":
    main()

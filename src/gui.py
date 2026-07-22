import queue
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from src.archive import create_pdf_archive
from src.converter import convert_documents
from src.models import ConversionResult


class ConverterApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Batch DOCX to PDF Converter")
        self.geometry("760x590")
        self.minsize(640, 500)
        self._set_window_icon()

        self.files: list[Path] = []
        self.file_items: dict[Path, str] = {}
        self.output_dir = tk.StringVar(
            value=str(Path.home() / "Documents" / "Converted PDFs")
        )
        self.create_zip = tk.BooleanVar(value=True)
        self.zip_name = tk.StringVar(value="converted_files.zip")
        self.progress_text = tk.StringVar(value="0%")
        self.status = tk.StringVar(value="Add Word documents to get started.")
        self.events: queue.Queue[tuple[str, object]] = queue.Queue()
        self._build_ui()

    def _set_window_icon(self) -> None:
        icon_path = Path(__file__).resolve().parent.parent / "assets" / "app.ico"
        if icon_path.is_file():
            try:
                self.iconbitmap(default=str(icon_path))
            except tk.TclError:
                pass

    def _build_ui(self) -> None:
        style = ttk.Style(self)
        style.configure("Title.TLabel", font=("Montserrat", 18, "bold"))
        style.configure("Section.TLabel", font=("Montserrat", 10, "bold"))
        style.configure("Convert.TButton", padding=(20, 8))

        container = ttk.Frame(self, padding=18)
        container.pack(fill="both", expand=True)
        container.columnconfigure(0, weight=1)
        container.rowconfigure(3, weight=1)

        ttk.Label(
            container, text="Batch DOCX → PDF Converter", style="Title.TLabel"
        ).grid(row=0, column=0, sticky="w", pady=(0, 16))
        ttk.Separator(container).grid(row=1, column=0, sticky="ew", pady=(0, 12))

        actions = ttk.Frame(container)
        actions.grid(row=2, column=0, sticky="w", pady=(0, 10))
        ttk.Button(actions, text="Add Files", command=self._add_files).pack(side="left")
        ttk.Button(actions, text="Remove", command=self._remove_selected).pack(
            side="left", padx=8
        )
        ttk.Button(actions, text="Clear", command=self._clear_files).pack(side="left")

        list_frame = ttk.Frame(container)
        list_frame.grid(row=3, column=0, sticky="nsew")
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)
        self.file_list = ttk.Treeview(
            list_frame,
            columns=("status", "file"),
            show="headings",
            selectmode="extended",
            height=9,
        )
        self.file_list.heading("status", text="")
        self.file_list.heading("file", text="Selected Word files")
        self.file_list.column(
            "status", width=48, minwidth=48, stretch=False, anchor="center"
        )
        self.file_list.column("file", width=600, anchor="w")
        self.file_list.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(
            list_frame, orient="vertical", command=self.file_list.yview
        )
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.file_list.configure(yscrollcommand=scrollbar.set)

        output_frame = ttk.LabelFrame(container, text="Output", padding=12)
        output_frame.grid(row=4, column=0, sticky="ew", pady=(14, 12))
        output_frame.columnconfigure(1, weight=1)
        ttk.Label(output_frame, text="Output Folder:").grid(
            row=0, column=0, sticky="w", padx=(0, 8)
        )
        ttk.Entry(output_frame, textvariable=self.output_dir).grid(
            row=0, column=1, sticky="ew"
        )
        ttk.Button(output_frame, text="Browse", command=self._choose_output_dir).grid(
            row=0, column=2, padx=(8, 0)
        )
        ttk.Checkbutton(
            output_frame,
            text="Create ZIP after conversion",
            variable=self.create_zip,
            command=self._toggle_zip_name,
        ).grid(row=1, column=0, columnspan=3, sticky="w", pady=(12, 8))
        ttk.Label(output_frame, text="ZIP Name:").grid(
            row=2, column=0, sticky="w", padx=(0, 8)
        )
        self.zip_entry = ttk.Entry(output_frame, textvariable=self.zip_name)
        self.zip_entry.grid(row=2, column=1, sticky="ew")

        progress_header = ttk.Frame(container)
        progress_header.grid(row=5, column=0, sticky="ew")
        progress_header.columnconfigure(0, weight=1)
        ttk.Label(progress_header, text="Progress:", style="Section.TLabel").grid(
            row=0, column=0, sticky="w"
        )
        ttk.Label(progress_header, textvariable=self.progress_text).grid(
            row=0, column=1, sticky="e"
        )
        self.progress = ttk.Progressbar(container, mode="determinate", maximum=100)
        self.progress.grid(row=6, column=0, sticky="ew", pady=(5, 10))

        ttk.Label(container, text="Status:", style="Section.TLabel").grid(
            row=7, column=0, sticky="w"
        )
        ttk.Label(container, textvariable=self.status).grid(
            row=8, column=0, sticky="w", pady=(3, 14)
        )
        self.convert_button = ttk.Button(
            container,
            text="Convert",
            style="Convert.TButton",
            command=self._start_conversion,
        )
        self.convert_button.grid(row=9, column=0, sticky="e")

    def _add_files(self) -> None:
        selected = filedialog.askopenfilenames(
            title="Choose Word documents",
            filetypes=[("Word documents", "*.docx *.doc"), ("All files", "*.*")],
        )
        known = set(self.files)
        for name in selected:
            path = Path(name)
            if path not in known:
                self.files.append(path)
                self.file_items[path] = self.file_list.insert(
                    "", tk.END, values=("✓", path.name)
                )
                known.add(path)
        self.status.set(f"{len(self.files)} file(s) selected.")

    def _remove_selected(self) -> None:
        selected_ids = set(self.file_list.selection())
        removed = [
            path for path, item_id in self.file_items.items() if item_id in selected_ids
        ]
        for path in removed:
            self.files.remove(path)
            self.file_list.delete(self.file_items.pop(path))
        self.status.set(f"{len(self.files)} file(s) selected.")

    def _clear_files(self) -> None:
        self.files.clear()
        self.file_items.clear()
        self.file_list.delete(*self.file_list.get_children())
        self.progress.configure(value=0)
        self.progress_text.set("0%")
        self.status.set("Add Word documents to get started.")

    def _choose_output_dir(self) -> None:
        selected = filedialog.askdirectory(
            title="Choose output folder", initialdir=self.output_dir.get()
        )
        if selected:
            self.output_dir.set(selected)

    def _toggle_zip_name(self) -> None:
        self.zip_entry.configure(
            state="normal" if self.create_zip.get() else "disabled"
        )

    def _validated_zip_name(self) -> str | None:
        name = self.zip_name.get().strip()
        if not name:
            messagebox.showwarning(
                "Missing ZIP name", "Enter a name for the ZIP archive."
            )
            return None
        if not name.lower().endswith(".zip"):
            name += ".zip"
        if Path(name).name != name or name in {".zip", "..zip"}:
            messagebox.showwarning(
                "Invalid ZIP name", "Enter a filename such as converted_files.zip."
            )
            return None
        self.zip_name.set(name)
        return name

    def _start_conversion(self) -> None:
        if not self.files:
            messagebox.showwarning("No files", "Choose at least one Word document.")
            return
        output_text = self.output_dir.get().strip()
        if not output_text:
            messagebox.showwarning("No output folder", "Choose an output folder.")
            return

        archive_name: str | None = None
        if self.create_zip.get():
            archive_name = self._validated_zip_name()
            if archive_name is None:
                return

        for item_id in self.file_items.values():
            filename = self.file_list.item(item_id, "values")[1]
            self.file_list.item(item_id, values=("…", filename))

        output_dir = Path(output_text).expanduser()
        self.convert_button.configure(state="disabled")
        self.progress.configure(value=0)
        self.progress_text.set("0%")
        self.status.set(f"Converted 0 of {len(self.files)} files...")
        threading.Thread(
            target=self._convert_worker,
            args=(list(self.files), output_dir, archive_name),
            daemon=True,
        ).start()
        self.after(100, self._process_events)

    def _convert_worker(
        self, files: list[Path], output_dir: Path, archive_name: str | None
    ) -> None:
        def report(result: ConversionResult, current: int, total: int) -> None:
            self.events.put(("progress", (result, current, total)))

        results = convert_documents(files, output_dir, report)
        successful = [
            result.output for result in results if result.succeeded and result.output
        ]
        archive: Path | None = None
        archive_error: str | None = None
        if archive_name and successful:
            try:
                archive = create_pdf_archive(successful, output_dir / archive_name)
            except Exception as exc:
                archive_error = str(exc)
        self.events.put(("done", (results, archive, archive_error, output_dir)))

    def _process_events(self) -> None:
        try:
            while True:
                event, payload = self.events.get_nowait()
                if event == "progress":
                    result, current, total = payload
                    percent = round(current / total * 100) if total else 0
                    self.progress.configure(value=percent)
                    self.progress_text.set(f"{percent}%")
                    item_id = self.file_items.get(result.source)
                    if item_id:
                        marker = "✓" if result.succeeded else "✕"
                        self.file_list.item(
                            item_id, values=(marker, result.source.name)
                        )
                    self.status.set(f"Converted {current} of {total} files...")
                elif event == "done":
                    self._finish_conversion(*payload)
                    return
        except queue.Empty:
            self.after(100, self._process_events)

    def _finish_conversion(
        self,
        results: list[ConversionResult],
        archive: Path | None,
        archive_error: str | None,
        output_dir: Path,
    ) -> None:
        self.convert_button.configure(state="normal")
        failures = [result for result in results if not result.succeeded]
        success_count = len(results) - len(failures)
        self.status.set(f"Finished: {success_count} converted, {len(failures)} failed.")

        details = [
            f"Converted {success_count} of {len(results)} file(s).",
            f"Output: {output_dir}",
        ]
        if archive:
            details.append(f"ZIP: {archive}")
        if archive_error:
            details.append(f"ZIP error: {archive_error}")
        if failures:
            details.append("\nFailed files:")
            details.extend(f"- {item.source.name}: {item.error}" for item in failures)
            messagebox.showwarning(
                "Conversion completed with errors", "\n".join(details)
            )
        else:
            messagebox.showinfo("Conversion complete", "\n".join(details))

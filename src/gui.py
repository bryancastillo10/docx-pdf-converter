import queue
import threading
import tkinter as tk
from math import ceil
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from src.archive import create_pdf_archive
from src.converter import convert_documents
from src.i18n import DEFAULT_LANGUAGE, LANGUAGE_OPTIONS, translate
from src.models import ConversionResult


class ConverterApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.geometry("760x640")
        self.minsize(640, 550)
        self._set_window_icon()
        self.header_image = self._load_header_image()

        self.language = DEFAULT_LANGUAGE
        self.language_choice = tk.StringVar(value="EN")
        self._status_key = "ready"
        self._status_values: dict[str, object] = {}
        self.files: list[Path] = []
        self.file_items: dict[Path, str] = {}
        self.output_dir = tk.StringVar(
            value=str(Path.home() / "Documents" / "Converted PDFs")
        )
        self.create_zip = tk.BooleanVar(value=True)
        self.zip_name = tk.StringVar(value="converted_files.zip")
        self.progress_text = tk.StringVar(value="0%")
        self.status = tk.StringVar()
        self.events: queue.Queue[tuple[str, object]] = queue.Queue()
        self._build_ui()
        self._apply_language()

    def _set_window_icon(self) -> None:
        icon_path = Path(__file__).resolve().parent.parent / "assets" / "logo.ico"
        if icon_path.is_file():
            try:
                self.iconbitmap(default=str(icon_path))
            except tk.TclError:
                pass

    def _load_header_image(self) -> tk.PhotoImage | None:
        image_path = (
            Path(__file__).resolve().parent.parent / "assets" / "main_app.png"
        )
        if not image_path.is_file():
            return None

        try:
            image = tk.PhotoImage(file=str(image_path))
        except tk.TclError:
            return None

        max_size = 76
        scale = max(1, ceil(max(image.width(), image.height()) / max_size))
        return image.subsample(scale, scale) if scale > 1 else image

    def _build_ui(self) -> None:
        style = ttk.Style(self)
        style.configure("Title.TLabel", font=("Montserrat", 18, "bold"))
        style.configure("Section.TLabel", font=("Montserrat", 10, "bold"))
        style.configure("Convert.TButton", padding=(20, 8))

        container = ttk.Frame(self, padding=18)
        container.pack(fill="both", expand=True)
        container.columnconfigure(0, weight=1)
        container.rowconfigure(3, weight=1)

        header = ttk.Frame(container)
        header.grid(row=0, column=0, sticky="ew", pady=(0, 16))
        header.columnconfigure(1, weight=1)
        self.header_image_label = ttk.Label(
            header, image=self.header_image if self.header_image else ""
        )
        self.header_image_label.grid(row=0, column=0, padx=(0, 14))
        self.title_label = ttk.Label(header, style="Title.TLabel")
        self.title_label.grid(row=0, column=1, sticky="w")
        self.language_label = ttk.Label(header)
        self.language_label.grid(row=0, column=2, padx=(12, 6))
        self.language_selector = ttk.Combobox(
            header,
            textvariable=self.language_choice,
            values=tuple(LANGUAGE_OPTIONS),
            state="readonly",
            width=7,
        )
        self.language_selector.grid(row=0, column=3)
        self.language_selector.bind("<<ComboboxSelected>>", self._change_language)
        ttk.Separator(container).grid(row=1, column=0, sticky="ew", pady=(0, 12))

        actions = ttk.Frame(container)
        actions.grid(row=2, column=0, sticky="w", pady=(0, 10))
        self.add_button = ttk.Button(actions, command=self._add_files)
        self.add_button.pack(side="left")
        self.remove_button = ttk.Button(actions, command=self._remove_selected)
        self.remove_button.pack(side="left", padx=8)
        self.clear_button = ttk.Button(actions, command=self._clear_files)
        self.clear_button.pack(side="left")

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

        self.output_frame = ttk.LabelFrame(container, padding=12)
        self.output_frame.grid(row=4, column=0, sticky="ew", pady=(14, 12))
        self.output_frame.columnconfigure(1, weight=1)
        self.output_folder_label = ttk.Label(self.output_frame)
        self.output_folder_label.grid(row=0, column=0, sticky="w", padx=(0, 8))
        ttk.Entry(self.output_frame, textvariable=self.output_dir).grid(
            row=0, column=1, sticky="ew"
        )
        self.browse_button = ttk.Button(
            self.output_frame, command=self._choose_output_dir
        )
        self.browse_button.grid(row=0, column=2, padx=(8, 0))
        self.zip_checkbutton = ttk.Checkbutton(
            self.output_frame,
            variable=self.create_zip,
            command=self._toggle_zip_name,
        )
        self.zip_checkbutton.grid(
            row=1, column=0, columnspan=3, sticky="w", pady=(12, 8)
        )
        self.zip_name_label = ttk.Label(self.output_frame)
        self.zip_name_label.grid(row=2, column=0, sticky="w", padx=(0, 8))
        self.zip_entry = ttk.Entry(self.output_frame, textvariable=self.zip_name)
        self.zip_entry.grid(row=2, column=1, sticky="ew")

        progress_header = ttk.Frame(container)
        progress_header.grid(row=5, column=0, sticky="ew")
        progress_header.columnconfigure(0, weight=1)
        self.progress_label = ttk.Label(progress_header, style="Section.TLabel")
        self.progress_label.grid(row=0, column=0, sticky="w")
        ttk.Label(progress_header, textvariable=self.progress_text).grid(
            row=0, column=1, sticky="e"
        )
        self.progress = ttk.Progressbar(container, mode="determinate", maximum=100)
        self.progress.grid(row=6, column=0, sticky="ew", pady=(5, 10))

        self.status_label = ttk.Label(container, style="Section.TLabel")
        self.status_label.grid(row=7, column=0, sticky="w")
        ttk.Label(container, textvariable=self.status).grid(
            row=8, column=0, sticky="w", pady=(3, 14)
        )
        self.convert_button = ttk.Button(
            container,
            style="Convert.TButton",
            command=self._start_conversion,
        )
        self.convert_button.grid(row=9, column=0, sticky="e")

    def _t(self, key: str, **values: object) -> str:
        return translate(self.language, key, **values)

    def _set_status(self, key: str, **values: object) -> None:
        self._status_key = key
        self._status_values = values
        self.status.set(self._t(key, **values))

    def _change_language(self, _event: object = None) -> None:
        self.language = LANGUAGE_OPTIONS.get(
            self.language_choice.get(), DEFAULT_LANGUAGE
        )
        self._apply_language()

    def _apply_language(self) -> None:
        self.title(self._t("window_title"))
        self.title_label.configure(text=self._t("app_title"))
        self.language_label.configure(text=self._t("language"))
        self.add_button.configure(text=self._t("add_files"))
        self.remove_button.configure(text=self._t("remove"))
        self.clear_button.configure(text=self._t("clear"))
        self.file_list.heading("file", text=self._t("selected_files"))
        self.output_frame.configure(text=self._t("output"))
        self.output_folder_label.configure(text=self._t("output_folder"))
        self.browse_button.configure(text=self._t("browse"))
        self.zip_checkbutton.configure(text=self._t("create_zip"))
        self.zip_name_label.configure(text=self._t("zip_name"))
        self.progress_label.configure(text=self._t("progress"))
        self.status_label.configure(text=self._t("status"))
        self.convert_button.configure(text=self._t("convert"))
        self.status.set(self._t(self._status_key, **self._status_values))

    def _add_files(self) -> None:
        selected = filedialog.askopenfilenames(
            title=self._t("choose_documents"),
            filetypes=[
                (self._t("office_documents"), "*.docx *.doc *.pptx *.ppt"),
                (self._t("word_documents"), "*.docx *.doc"),
                (self._t("powerpoint_presentations"), "*.pptx *.ppt"),
                (self._t("all_files"), "*.*"),
            ],
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
        self._set_status("files_selected", count=len(self.files))

    def _remove_selected(self) -> None:
        selected_ids = set(self.file_list.selection())
        removed = [
            path for path, item_id in self.file_items.items() if item_id in selected_ids
        ]
        for path in removed:
            self.files.remove(path)
            self.file_list.delete(self.file_items.pop(path))
        self._set_status("files_selected", count=len(self.files))

    def _clear_files(self) -> None:
        self.files.clear()
        self.file_items.clear()
        self.file_list.delete(*self.file_list.get_children())
        self.progress.configure(value=0)
        self.progress_text.set("0%")
        self._set_status("ready")

    def _choose_output_dir(self) -> None:
        selected = filedialog.askdirectory(
            title=self._t("choose_output"), initialdir=self.output_dir.get()
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
                self._t("missing_zip_title"), self._t("missing_zip_message")
            )
            return None
        if not name.lower().endswith(".zip"):
            name += ".zip"
        if Path(name).name != name or name in {".zip", "..zip"}:
            messagebox.showwarning(
                self._t("invalid_zip_title"), self._t("invalid_zip_message")
            )
            return None
        self.zip_name.set(name)
        return name

    def _start_conversion(self) -> None:
        if not self.files:
            messagebox.showwarning(
                self._t("no_files_title"), self._t("no_files_message")
            )
            return
        output_text = self.output_dir.get().strip()
        if not output_text:
            messagebox.showwarning(
                self._t("no_output_title"), self._t("no_output_message")
            )
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
        self._set_status("converting", current=0, total=len(self.files))
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
                    self._set_status("converting", current=current, total=total)
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
        self._set_status("finished", success=success_count, failed=len(failures))

        details = [
            self._t("result_summary", success=success_count, total=len(results)),
            self._t("result_output", path=output_dir),
        ]
        if archive:
            details.append(self._t("result_zip", path=archive))
        if archive_error:
            details.append(self._t("zip_error", error=archive_error))
        if failures:
            details.append(f"\n{self._t('failed_files')}")
            details.extend(f"- {item.source.name}: {item.error}" for item in failures)
            messagebox.showwarning(
                self._t("completed_errors_title"), "\n".join(details)
            )
        else:
            messagebox.showinfo(self._t("completed_title"), "\n".join(details))

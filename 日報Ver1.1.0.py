from __future__ import annotations

import json
import tkinter as tk
from dataclasses import dataclass
from datetime import date, datetime, time, timedelta
from pathlib import Path
from tkinter import messagebox, ttk

import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.styles.borders import Border, Side


APP_VERSION = "2.0.0"
CONFIG_NAME = "config.json"
LEGACY_CONFIG_NAME = "excelpath.xlsx"
DEMO_DIR_NAME = "demo_data"
REPORT_HEADERS = ["日付", "名前", "工番", "作業区分", "作業番号", "開始", "終了", "所要時間", "残業", "夜勤"]
MONTHS = [str(month) for month in range(4, 13)] + [str(month) for month in range(1, 4)]
DAYS = [str(day) for day in range(1, 32)]
HOURS = [str(hour) for hour in list(range(8, 24)) + list(range(0, 8))]
MINUTES = ["0", "10", "20", "30", "40", "50"]
END_MINUTES = MINUTES + ["59"]
BREAK_MINUTES = ["0", "10", "20", "30", "40", "50", "60", "70", "80", "90", "100", "110", "120"]
CLASS_COLUMNS = {
    0: ("C", "N", "W", "B6"),
    1: ("D", "O", "X", "B7"),
    2: ("E", "P", "Y", "B8"),
    3: ("F", "Q", "Z", "B9"),
    4: ("G", "N", "W", "B10"),
    5: ("H", "P", "Y", "B11"),
    6: ("I", "R", "Z", "B12"),
    7: ("J", "S", "AA", "B13"),
    8: ("K", "T", "AB", "B14"),
}


class AppError(Exception):
    pass


class ValidationError(AppError):
    pass


class WorkbookError(AppError):
    pass


@dataclass(frozen=True)
class AppConfig:
    master_path: Path
    report_path: Path
    overtime_path: Path
    factory_mode: int = 0
    class_mode: int = 0
    default_name_index: int | None = None
    append_fast: bool = False
    demo_mode: bool = False
    source: str = CONFIG_NAME


@dataclass(frozen=True)
class MasterData:
    materials: list[str]
    material_search: list[str]
    work_classes: list[str]
    work_numbers: list[str]
    names: list[str]


@dataclass(frozen=True)
class TimeResult:
    total_minutes: int
    break_minutes: int
    nightshift_minutes: int


@dataclass
class WorkEntry:
    work_date: date
    name: str
    material_number: str
    work_class: str
    work_number: str
    start_time: time
    end_time: time
    work_class_detail: str = ""
    manual_break_minutes: int | None = None
    result: TimeResult | None = None

    @property
    def output_work_class(self) -> str:
        return f"{self.work_class}{self.work_class_detail}"


def load_config(base_dir: Path) -> AppConfig:
    json_path = base_dir / CONFIG_NAME
    if json_path.exists():
        try:
            raw = json.loads(json_path.read_text(encoding="utf-8"))
            return AppConfig(
                master_path=resolve_path(base_dir, raw["master_path"]),
                report_path=resolve_path(base_dir, raw["report_path"]),
                overtime_path=resolve_path(base_dir, raw["overtime_path"]),
                factory_mode=int(raw.get("factory_mode", 0)),
                class_mode=int(raw.get("class_mode", 0)),
                default_name_index=optional_int(raw.get("default_name_index")),
                append_fast=bool(raw.get("append_fast", False)),
                source=str(json_path),
            )
        except (OSError, json.JSONDecodeError, KeyError, TypeError, ValueError) as exc:
            raise AppError(f"{CONFIG_NAME} を読めません: {exc}") from exc

    legacy_path = base_dir / LEGACY_CONFIG_NAME
    if legacy_path.exists():
        return load_legacy_config(legacy_path)

    return create_demo_workbooks(base_dir / DEMO_DIR_NAME)


def resolve_path(base_dir: Path, value: object) -> Path:
    if value is None or str(value).strip() == "":
        raise AppError("Excel パスが空です。")
    path = Path(str(value))
    return path if path.is_absolute() else (base_dir / path).resolve()


def optional_int(value: object) -> int | None:
    return None if value in (None, "") else int(value)


def load_legacy_config(path: Path) -> AppConfig:
    try:
        workbook = openpyxl.load_workbook(path, data_only=True)
        sheet = workbook.worksheets[0]
        class_mode = int(sheet["D3"].value or 0)
        report_cell = CLASS_COLUMNS.get(class_mode, CLASS_COLUMNS[0])[3]
        config = AppConfig(
            master_path=resolve_path(path.parent, sheet["B1"].value),
            report_path=resolve_path(path.parent, sheet[report_cell].value),
            overtime_path=resolve_path(path.parent, sheet["B3"].value),
            factory_mode=int(sheet["D1"].value or 0),
            class_mode=class_mode,
            default_name_index=optional_int(sheet["D4"].value),
            append_fast=bool(int(sheet["D2"].value or 0)),
            source=str(path),
        )
        workbook.close()
        return config
    except (IndexError, OSError, TypeError, ValueError) as exc:
        raise AppError(f"{LEGACY_CONFIG_NAME} を読めません: {exc}") from exc


def create_demo_workbooks(demo_dir: Path) -> AppConfig:
    demo_dir.mkdir(exist_ok=True)
    master_path = demo_dir / "demo_master.xlsx"
    report_path = demo_dir / "demo_daily_report.xlsx"
    overtime_path = demo_dir / "demo_overtime.xlsx"

    if not master_path.exists():
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.title = "master"
        sheet["A1"] = "工番"
        sheet["C1"] = "名前"
        sheet["N1"] = "作業区分"
        sheet["W1"] = "作業番号"
        for row, value in enumerate(["MAT0001 デモ製品A", "MAT0002 デモ製品B", "MAT0003 治具"], start=2):
            sheet.cell(row, 1).value = value
        for row, value in enumerate(["山田", "佐藤", "Demo User"], start=2):
            sheet.cell(row, 3).value = value
        for row, value in enumerate(["組立", "検査", "段取"], start=2):
            sheet.cell(row, 14).value = value
        for row, value in enumerate(["WORK-10", "WORK-20", "WORK-30"], start=2):
            sheet.cell(row, 23).value = value
        style_headers(sheet, ["A1", "C1", "N1", "W1"])
        workbook.save(master_path)

    if not report_path.exists():
        create_table_workbook(report_path, REPORT_HEADERS)
    if not overtime_path.exists():
        create_table_workbook(overtime_path, ["日付", "名前"])

    return AppConfig(master_path, report_path, overtime_path, demo_mode=True, source="自動生成デモ Excel")


def create_table_workbook(path: Path, headers: list[str]) -> None:
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "data"
    for column, header in enumerate(headers, start=1):
        sheet.cell(1, column).value = header
    style_headers(sheet, [sheet.cell(1, column).coordinate for column in range(1, len(headers) + 1)])
    workbook.save(path)


def style_headers(sheet, cells: list[str]) -> None:
    for cell_name in cells:
        cell = sheet[cell_name]
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="DDEBF7")


def validate_workbook_path(path: Path, label: str) -> None:
    if not path.exists():
        raise WorkbookError(f"{label} が見つかりません: {path}")


def open_workbook(path: Path):
    validate_workbook_path(path, "Excel ファイル")
    try:
        return openpyxl.load_workbook(path, keep_vba=path.suffix.lower() == ".xlsm")
    except (OSError, ValueError) as exc:
        raise WorkbookError(f"Excel ファイルを開けません: {path}\n{exc}") from exc


def values_until_blank(sheet, column: str) -> list[str]:
    values: list[str] = []
    for row in range(2, sheet.max_row + 1):
        value = sheet[f"{column}{row}"].value
        if value in (None, ""):
            continue
        values.append(str(value))
    return values


def load_master_data(config: AppConfig) -> MasterData:
    workbook = open_workbook(config.master_path)
    try:
        sheet = workbook.worksheets[0]
        name_column, class_column, work_column, _ = CLASS_COLUMNS.get(config.class_mode, CLASS_COLUMNS[0])
        material_search = values_until_blank(sheet, "A")
        materials = [value[:7] for value in material_search]
        data = MasterData(
            materials=materials,
            material_search=material_search,
            work_classes=values_until_blank(sheet, class_column),
            work_numbers=values_until_blank(sheet, work_column),
            names=values_until_blank(sheet, name_column),
        )
        if not all([data.materials, data.work_classes, data.work_numbers, data.names]):
            raise WorkbookError("マスタ Excel の工番、作業区分、作業番号、名前が不足しています。")
        return data
    finally:
        workbook.close()


def first_empty_report_row(sheet) -> int:
    for row in range(2, max(sheet.max_row, 2) + 2):
        if sheet.cell(row, 5).value in (None, ""):
            return row
    return sheet.max_row + 1


def entry_to_row(entry: WorkEntry, overtime: int | str, nightshift: int | str) -> list[object]:
    if entry.result is None:
        raise ValidationError("計算されていない作業行があります。")
    return [
        entry.work_date,
        entry.name,
        entry.material_number,
        entry.output_work_class,
        entry.work_number,
        entry.start_time,
        entry.end_time,
        entry.result.total_minutes,
        overtime,
        nightshift,
    ]


def append_entries(config: AppConfig, entries: list[WorkEntry]) -> None:
    workbook = open_workbook(config.report_path)
    try:
        sheet = workbook.worksheets[0]
        row = first_empty_report_row(sheet)
        total_overtime = overtime_minutes(entries)
        total_nightshift = sum(entry.result.nightshift_minutes for entry in entries if entry.result)
        for index, entry in enumerate(entries):
            report_row = entry_to_row(
                entry,
                total_overtime if index == len(entries) - 1 and total_overtime else "",
                total_nightshift if index == len(entries) - 1 and total_nightshift else "",
            )
            write_report_row(sheet, row + index, report_row)
        save_workbook(workbook, config.report_path)
    finally:
        workbook.close()


def update_entries(config: AppConfig, entries: list[WorkEntry]) -> None:
    if not entries:
        raise ValidationError("修正する作業行がありません。")
    workbook = open_workbook(config.report_path)
    try:
        sheet = workbook.worksheets[0]
        rows = find_report_rows(sheet, entries[0].name, entries[0].work_date)
        if not rows:
            raise WorkbookError("同じ日付と名前の登録済みデータが見つかりません。")
        for row in reversed(rows[len(entries):]):
            sheet.delete_rows(row)
        total_overtime = overtime_minutes(entries)
        total_nightshift = sum(entry.result.nightshift_minutes for entry in entries if entry.result)
        while len(rows) < len(entries):
            insert_at = rows[-1] + 1 if rows else first_empty_report_row(sheet)
            sheet.insert_rows(insert_at)
            rows.append(insert_at)
        for index, entry in enumerate(entries):
            write_report_row(
                sheet,
                rows[index],
                entry_to_row(
                    entry,
                    total_overtime if index == len(entries) - 1 and total_overtime else "",
                    total_nightshift if index == len(entries) - 1 and total_nightshift else "",
                ),
            )
        save_workbook(workbook, config.report_path)
    finally:
        workbook.close()


def find_report_rows(sheet, name: str, work_date: date) -> list[int]:
    rows: list[int] = []
    for row in range(2, sheet.max_row + 1):
        raw_date = sheet.cell(row, 1).value
        raw_name = sheet.cell(row, 2).value
        saved_date = raw_date.date() if isinstance(raw_date, datetime) else raw_date
        if saved_date == work_date and raw_name == name:
            rows.append(row)
    return rows


def load_existing_entries(config: AppConfig, name: str, work_date: date) -> list[WorkEntry]:
    workbook = open_workbook(config.report_path)
    try:
        sheet = workbook.worksheets[0]
        entries = []
        for row in find_report_rows(sheet, name, work_date):
            start_at = sheet.cell(row, 6).value
            end_at = sheet.cell(row, 7).value
            if isinstance(start_at, datetime):
                start_at = start_at.time()
            if isinstance(end_at, datetime):
                end_at = end_at.time()
            if not isinstance(start_at, time) or not isinstance(end_at, time):
                continue
            entries.append(
                WorkEntry(
                    work_date=work_date,
                    name=name,
                    material_number=str(sheet.cell(row, 3).value or ""),
                    work_class=str(sheet.cell(row, 4).value or ""),
                    work_number=str(sheet.cell(row, 5).value or ""),
                    start_time=start_at,
                    end_time=end_at,
                )
            )
        return entries
    finally:
        workbook.close()


def write_report_row(sheet, row: int, values: list[object]) -> None:
    border = Border(*(Side(style="thin", color="000000"),) * 4)
    for column, value in enumerate(values, start=1):
        cell = sheet.cell(row, column)
        cell.value = value
        cell.border = border
    sheet.cell(row, 1).number_format = "m月d日"


def write_overtime_check(config: AppConfig, work_date: date, name: str) -> None:
    workbook = open_workbook(config.overtime_path)
    try:
        sheet = workbook.worksheets[0]
        row = max(sheet.max_row + 1, 2)
        sheet.cell(row, 1).value = work_date
        sheet.cell(row, 2).value = name
        save_workbook(workbook, config.overtime_path)
    finally:
        workbook.close()


def save_workbook(workbook, path: Path) -> None:
    try:
        workbook.save(path)
    except PermissionError as exc:
        raise WorkbookError(f"保存できません。Excel で開いている可能性があります: {path}") from exc
    except OSError as exc:
        raise WorkbookError(f"保存できません: {path}\n{exc}") from exc


def calculate_time(entry: WorkEntry, factory_mode: int) -> TimeResult:
    start_at = datetime.combine(entry.work_date, entry.start_time)
    end_at = datetime.combine(entry.work_date, entry.end_time)
    if end_at < start_at:
        raise ValidationError("終了時間は開始時間より後にしてください。")
    corrected_end = end_at + timedelta(minutes=1) if entry.end_time == time(23, 59) else end_at
    span = int((corrected_end - start_at).total_seconds() // 60)
    if span <= 0:
        raise ValidationError("所要時間が 0 分です。")
    break_minutes = entry.manual_break_minutes
    if break_minutes is None:
        break_minutes = automatic_break_minutes(entry.work_date, start_at, corrected_end, factory_mode)
    if break_minutes < 0 or break_minutes >= span:
        raise ValidationError("休憩時間が作業時間の範囲外です。")
    return TimeResult(span - break_minutes, break_minutes, nightshift_minutes(entry.work_date, start_at, corrected_end))


def automatic_break_minutes(work_date: date, start_at: datetime, end_at: datetime, factory_mode: int) -> int:
    breaks = {
        0: [(time(10, 10), time(10, 20)), (time(13), time(14)), (time(15, 10), time(15, 20))],
        1: [(time(10), time(10, 15)), (time(13), time(14)), (time(15), time(15, 15))],
    }.get(factory_mode, [])
    return sum(overlap_minutes(start_at, end_at, datetime.combine(work_date, first), datetime.combine(work_date, second)) for first, second in breaks)


def nightshift_minutes(work_date: date, start_at: datetime, end_at: datetime) -> int:
    morning = overlap_minutes(start_at, end_at, datetime.combine(work_date, time()), datetime.combine(work_date, time(5)))
    night = overlap_minutes(start_at, end_at, datetime.combine(work_date, time(22)), datetime.combine(work_date + timedelta(days=1), time()))
    return morning + night


def overlap_minutes(left_start: datetime, left_end: datetime, right_start: datetime, right_end: datetime) -> int:
    return max(int((min(left_end, right_end) - max(left_start, right_start)).total_seconds() // 60), 0)


def overtime_minutes(entries: list[WorkEntry]) -> int:
    return max(sum(entry.result.total_minutes for entry in entries if entry.result) - 480, 0)


class EntryRow:
    def __init__(self, app: KousuApp, index: int, source: WorkEntry | None = None):
        self.app = app
        self.frame = ttk.Frame(app.rows_frame)
        self.month = tk.StringVar(value=str(source.work_date.month if source else app.today.month))
        self.day = tk.StringVar(value=str(source.work_date.day if source else app.today.day))
        self.material = tk.StringVar(value=source.material_number if source else "")
        self.work_class = tk.StringVar(value=source.work_class if source else "")
        self.detail = tk.StringVar()
        self.work_number = tk.StringVar(value=source.work_number if source else "")
        self.start_hour = tk.StringVar(value=str(source.start_time.hour if source else 8))
        self.start_minute = tk.StringVar(value=str(source.start_time.minute if source else (0 if app.config.factory_mode else 10)))
        self.end_hour = tk.StringVar(value=str(source.end_time.hour if source else ""))
        self.end_minute = tk.StringVar(value=str(source.end_time.minute if source else ""))
        self.break_minutes = tk.StringVar(value="0")
        self.total = tk.StringVar(value="-")
        self.check = tk.BooleanVar(value=False)
        self.widgets: list[tk.Widget] = []
        self._build()
        self.refresh_index(index)
        if source:
            self.recalculate(silent=True)

    def _build(self) -> None:
        values = self.app.master
        self.index_label = ttk.Label(self.frame, width=3)
        self.index_label.grid(row=0, column=0, padx=2, pady=3)
        self.widgets.append(ttk.Checkbutton(self.frame, variable=self.check))
        self.widgets[-1].grid(row=0, column=1, padx=2)
        for column, variable, options, width in [
            (2, self.month, MONTHS, 3),
            (3, self.day, DAYS, 3),
            (4, self.material, values.materials, 10),
            (5, self.work_class, values.work_classes, 12),
            (7, self.work_number, values.work_numbers, 12),
            (8, self.start_hour, HOURS, 3),
            (9, self.start_minute, MINUTES, 3),
            (10, self.end_hour, HOURS, 3),
            (11, self.end_minute, END_MINUTES, 3),
            (12, self.break_minutes, BREAK_MINUTES, 4),
        ]:
            state = "readonly"
            widget = ttk.Combobox(self.frame, textvariable=variable, values=options, width=width, state=state)
            widget.grid(row=0, column=column, padx=2, pady=3)
            widget.bind("<<ComboboxSelected>>", lambda _event: self.recalculate(silent=True))
            self.widgets.append(widget)
        detail = ttk.Entry(self.frame, textvariable=self.detail, width=11)
        detail.grid(row=0, column=6, padx=2)
        self.widgets.append(detail)
        ttk.Label(self.frame, textvariable=self.total, width=10).grid(row=0, column=13, padx=2)
        ttk.Button(self.frame, text="Copy", command=self.copy_previous, width=6).grid(row=0, column=14, padx=2)
        ttk.Button(self.frame, text="Clear", command=self.clear, width=6).grid(row=0, column=15, padx=2)
        ttk.Button(self.frame, text="Remove", command=lambda: self.app.remove_row(self), width=8).grid(row=0, column=16, padx=2)
        self.update_break_state()

    def refresh_index(self, index: int) -> None:
        self.index = index
        self.index_label.configure(text=str(index + 1))
        self.frame.grid(row=index, column=0, sticky="ew")

    def update_break_state(self) -> None:
        break_widget = self.widgets[-1]
        break_widget.configure(state="readonly" if self.app.manual_break.get() else "disabled")

    def clear(self) -> None:
        for variable in [self.material, self.work_class, self.detail, self.work_number, self.end_hour, self.end_minute]:
            variable.set("")
        self.break_minutes.set("0")
        self.check.set(False)
        self.total.set("-")

    def copy_previous(self) -> None:
        if self.index == 0:
            return
        previous = self.app.rows[self.index - 1]
        self.material.set(previous.material.get())
        self.work_class.set(previous.work_class.get())
        self.detail.set(previous.detail.get())
        self.work_number.set(previous.work_number.get())
        self.recalculate(silent=True)

    def set_material(self, material: str) -> None:
        self.material.set(material)

    def has_any_work(self) -> bool:
        return any([self.material.get(), self.work_class.get(), self.work_number.get(), self.end_hour.get(), self.end_minute.get()])

    def make_entry(self) -> WorkEntry:
        if not self.has_any_work():
            raise ValidationError("")
        if not self.app.name.get():
            raise ValidationError("名前を選択してください。")
        required = {
            "作業区分": self.work_class.get(),
            "作業番号": self.work_number.get(),
            "終了時": self.end_hour.get(),
            "終了分": self.end_minute.get(),
        }
        missing = [label for label, value in required.items() if value == ""]
        if missing:
            raise ValidationError(f"{self.index + 1} 行目が未入力です: {', '.join(missing)}")
        try:
            work_date = date(self.app.today.year, int(self.month.get()), int(self.day.get()))
            entry = WorkEntry(
                work_date=work_date,
                name=self.app.name.get(),
                material_number=self.material.get(),
                work_class=self.work_class.get(),
                work_class_detail=self.detail.get(),
                work_number=self.work_number.get(),
                start_time=time(int(self.start_hour.get()), int(self.start_minute.get())),
                end_time=time(int(self.end_hour.get()), int(self.end_minute.get())),
                manual_break_minutes=int(self.break_minutes.get()) if self.app.manual_break.get() else None,
            )
        except ValueError as exc:
            raise ValidationError(f"{self.index + 1} 行目の日付または時刻が不正です。") from exc
        entry.result = calculate_time(entry, self.app.config.factory_mode)
        return entry

    def recalculate(self, silent: bool = False) -> None:
        if not self.end_hour.get() or not self.end_minute.get():
            self.total.set("-")
            self.app.refresh_totals()
            return
        try:
            result = self.make_entry().result
            self.total.set(f"{result.total_minutes} 分")
        except ValidationError as exc:
            self.total.set("要確認")
            if not silent and str(exc):
                self.app.set_status(str(exc), error=True)
        self.app.refresh_totals()


class KousuApp:
    def __init__(self, root: tk.Tk, config: AppConfig, master: MasterData):
        self.root = root
        self.config = config
        self.master = master
        self.today = date.today()
        self.rows: list[EntryRow] = []
        self.name = tk.StringVar()
        self.search = tk.StringVar()
        self.manual_break = tk.BooleanVar(value=False)
        self.status = tk.StringVar()
        self.total = tk.StringVar(value="合計 0 分")
        self.overtime = tk.StringVar(value="残業 0 分")
        self.nightshift = tk.StringVar(value="夜勤 0 分")
        self._build()
        self.choose_default_name()
        for _ in range(4):
            self.add_row()
        self.set_status(self.startup_message())

    def _build(self) -> None:
        title = "関東工場 " if self.config.factory_mode == 1 else ""
        self.root.title(f"{title}工数入力 {APP_VERSION}")
        self.root.geometry("1320x720")
        self.root.minsize(1120, 620)
        ttk.Style().configure("Heading.TLabel", font=("Yu Gothic UI", 11, "bold"))

        toolbar = ttk.Frame(self.root, padding=10)
        toolbar.pack(fill="x")
        ttk.Label(toolbar, text="名前", style="Heading.TLabel").pack(side="left")
        name_box = ttk.Combobox(toolbar, textvariable=self.name, values=self.master.names, state="readonly", width=14)
        name_box.pack(side="left", padx=(6, 14))
        name_box.bind("<<ComboboxSelected>>", lambda _event: self.refresh_totals())
        ttk.Checkbutton(toolbar, text="休憩手入力", variable=self.manual_break, command=self.toggle_breaks).pack(side="left")
        ttk.Button(toolbar, text="1日一括", command=self.one_day).pack(side="left", padx=4)
        ttk.Button(toolbar, text="行追加", command=self.add_row).pack(side="left", padx=4)
        ttk.Button(toolbar, text="全削除", command=self.clear_all).pack(side="left", padx=4)
        ttk.Button(toolbar, text="修正読込", command=self.load_for_update).pack(side="left", padx=(18, 4))
        ttk.Button(toolbar, text="修正保存", command=self.save_update).pack(side="left", padx=4)
        ttk.Button(toolbar, text="登録", command=self.save_new).pack(side="right", padx=4)

        totals = ttk.Frame(self.root, padding=(10, 0, 10, 8))
        totals.pack(fill="x")
        for variable in [self.total, self.overtime, self.nightshift]:
            ttk.Label(totals, textvariable=variable, style="Heading.TLabel").pack(side="left", padx=(0, 24))
        target = "DEMO" if self.config.demo_mode else "保存先"
        ttk.Label(totals, text=f"{target}: {self.config.report_path}").pack(side="left")

        search = ttk.Labelframe(self.root, text="工番検索", padding=8)
        search.pack(fill="x", padx=10, pady=(0, 8))
        ttk.Entry(search, textvariable=self.search, width=28).pack(side="left")
        ttk.Button(search, text="検索", command=self.search_materials).pack(side="left", padx=4)
        ttk.Button(search, text="選択行へ挿入", command=self.insert_material).pack(side="left", padx=4)
        self.search_results = tk.Listbox(search, height=3, width=55)
        self.search_results.pack(side="left", padx=8)

        table = ttk.Frame(self.root, padding=(10, 0))
        table.pack(fill="both", expand=True)
        headers = ["#", "対象", "月", "日", "工番", "作業区分", "詳細", "作業番号", "開始時", "分", "終了時", "分", "休憩", "所要", "", "", ""]
        header = ttk.Frame(table)
        header.pack(fill="x")
        for column, text in enumerate(headers):
            ttk.Label(header, text=text, style="Heading.TLabel").grid(row=0, column=column, padx=2, sticky="w")
        canvas = tk.Canvas(table, highlightthickness=0)
        scrollbar = ttk.Scrollbar(table, orient="vertical", command=canvas.yview)
        self.rows_frame = ttk.Frame(canvas)
        self.rows_frame.bind("<Configure>", lambda _event: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self.rows_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        status = ttk.Label(self.root, textvariable=self.status, padding=10)
        status.pack(fill="x")

    def startup_message(self) -> str:
        if self.config.demo_mode:
            return "設定が無いためデモ Excel で起動しました。入力と保存を試せます。"
        return f"設定読込: {self.config.source}"

    def choose_default_name(self) -> None:
        if self.config.default_name_index is not None and 0 <= self.config.default_name_index < len(self.master.names):
            self.name.set(self.master.names[self.config.default_name_index])

    def set_status(self, text: str, error: bool = False) -> None:
        self.status.set(("エラー: " if error else "") + text)

    def add_row(self, source: WorkEntry | None = None) -> EntryRow:
        row = EntryRow(self, len(self.rows), source)
        self.rows.append(row)
        self.refresh_totals()
        return row

    def remove_row(self, row: EntryRow) -> None:
        if len(self.rows) == 1:
            row.clear()
            return
        row.frame.destroy()
        self.rows.remove(row)
        for index, existing in enumerate(self.rows):
            existing.refresh_index(index)
        self.refresh_totals()

    def toggle_breaks(self) -> None:
        for row in self.rows:
            row.update_break_state()
            row.recalculate(silent=True)
        self.set_status("休憩時間を手入力できます。" if self.manual_break.get() else "標準休憩を自動計算します。")

    def one_day(self) -> None:
        first = self.rows[0]
        first.start_hour.set("8")
        first.start_minute.set("0" if self.config.factory_mode else "10")
        first.end_hour.set("17")
        first.end_minute.set("30")
        first.recalculate()

    def clear_all(self) -> None:
        for row in list(self.rows[1:]):
            self.remove_row(row)
        self.rows[0].clear()
        self.search.set("")
        self.search_results.delete(0, tk.END)
        self.refresh_totals()
        self.set_status("入力を削除しました。")

    def search_materials(self) -> None:
        query = self.search.get().strip()
        self.search_results.delete(0, tk.END)
        for value in self.master.material_search:
            if query in value:
                self.search_results.insert(tk.END, value)
        self.set_status("工番を選び、対象チェック行へ挿入してください。")

    def insert_material(self) -> None:
        selection = self.search_results.curselection()
        if not selection:
            self.set_status("検索結果を選択してください。", error=True)
            return
        material = self.search_results.get(selection[0])[:7]
        targets = [row for row in self.rows if row.check.get()]
        if not targets:
            self.set_status("挿入対象の行にチェックを入れてください。", error=True)
            return
        for row in targets:
            row.set_material(material)
        self.set_status(f"{len(targets)} 行へ工番 {material} を挿入しました。")

    def valid_entries(self) -> list[WorkEntry]:
        entries: list[WorkEntry] = []
        for row in self.rows:
            try:
                entries.append(row.make_entry())
            except ValidationError as exc:
                if str(exc):
                    raise
        if not entries:
            raise ValidationError("登録する作業行を入力してください。")
        return entries

    def refresh_totals(self) -> None:
        entries = []
        for row in self.rows:
            try:
                entries.append(row.make_entry())
            except ValidationError:
                pass
        total = sum(entry.result.total_minutes for entry in entries if entry.result)
        night = sum(entry.result.nightshift_minutes for entry in entries if entry.result)
        self.total.set(f"合計 {total} 分")
        self.overtime.set(f"残業 {max(total - 480, 0)} 分")
        self.nightshift.set(f"夜勤 {night} 分")

    def save_new(self) -> None:
        self._save(update=False)

    def save_update(self) -> None:
        self._save(update=True)

    def _save(self, update: bool) -> None:
        try:
            entries = self.valid_entries()
            if sum(entry.result.total_minutes for entry in entries if entry.result) < 480:
                if not messagebox.askyesno("確認", "合計時間が 480 分未満です。保存しますか？"):
                    return
            if update:
                update_entries(self.config, entries)
                action = "修正保存"
            else:
                append_entries(self.config, entries)
                action = "登録"
            if overtime_minutes(entries):
                write_overtime_check(self.config, entries[0].work_date, entries[0].name)
            self.set_status(f"{action}しました: {len(entries)} 行")
            messagebox.showinfo("確認", f"{action}しました。")
        except AppError as exc:
            self.set_status(str(exc), error=True)
            messagebox.showerror("保存エラー", str(exc))

    def load_for_update(self) -> None:
        try:
            if not self.name.get():
                raise ValidationError("修正読込には名前を選択してください。")
            first = self.rows[0]
            work_date = date(self.today.year, int(first.month.get()), int(first.day.get()))
            entries = load_existing_entries(self.config, self.name.get(), work_date)
            if not entries:
                raise WorkbookError("同じ日付と名前の登録済みデータが見つかりません。")
            for row in list(self.rows):
                row.frame.destroy()
            self.rows.clear()
            for entry in entries:
                entry.work_class, entry.work_class_detail = split_work_class(entry.output_work_class, self.master.work_classes)
                self.add_row(entry)
            self.set_status(f"修正対象を読み込みました: {len(entries)} 行")
        except (AppError, ValueError) as exc:
            self.set_status(str(exc), error=True)
            messagebox.showerror("修正読込エラー", str(exc))


def split_work_class(value: str, work_classes: list[str]) -> tuple[str, str]:
    candidates = sorted((work_class for work_class in work_classes if value.startswith(work_class)), key=len, reverse=True)
    if not candidates:
        return value, ""
    work_class = candidates[0]
    return work_class, value[len(work_class):]


def main() -> None:
    base_dir = Path(__file__).resolve().parent
    root = tk.Tk()
    root.withdraw()
    try:
        config = load_config(base_dir)
        master = load_master_data(config)
    except AppError as exc:
        messagebox.showerror("起動エラー", str(exc))
        root.destroy()
        return
    root.deiconify()
    KousuApp(root, config, master)
    root.mainloop()


if __name__ == "__main__":
    main()

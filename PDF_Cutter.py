import os
import re
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from PyPDF2 import PdfReader, PdfWriter


def safe_filename(name: str) -> str:
    bad = r'\\/:*?"<>|'
    for ch in bad:
        name = name.replace(ch, "_")
    name = name.strip()
    return name if name else "untitled"


class SplitRow:
    def __init__(self, parent, row_index, grid_row, on_change, confirm_switch, default_name=None):
        self.row_index = row_index
        self.on_change = on_change
        self.confirm_switch = confirm_switch

        self.mode_use_end = True
        default = default_name if default_name else f"Filename_{row_index + 1}"
        self.filename_var = tk.StringVar(value=default)
        self.start_var = tk.StringVar()
        self.end_var = tk.StringVar()
        self.count_var = tk.StringVar()

        self.filename_entry = ttk.Entry(parent, textvariable=self.filename_var, width=28, justify="left")
        self.start_entry = ttk.Entry(parent, textvariable=self.start_var, width=4, justify="center")
        self.end_frame = ttk.Frame(parent)
        self.count_frame = ttk.Frame(parent)
        self.end_entry = ttk.Entry(self.end_frame, textvariable=self.end_var, width=4, justify="center")
        self.count_entry = ttk.Entry(self.count_frame, textvariable=self.count_var, width=4, justify="center")
        self.end_block = tk.Label(self.end_frame, text="X", fg="#9a9a9a")
        self.count_block = tk.Label(self.count_frame, text="X", fg="#9a9a9a")
        self.switch_btn = ttk.Button(parent, text="<>", width=3, command=self.toggle_mode)
        self.vline_1 = tk.Frame(parent, bg="#d0d0d0", width=1)
        self.vline_2 = tk.Frame(parent, bg="#d0d0d0", width=1)
        self.vline_3 = tk.Frame(parent, bg="#d0d0d0", width=1)
        self.vline_4 = tk.Frame(parent, bg="#d0d0d0", width=1)

        self.filename_entry.grid(row=grid_row, column=0, padx=4, pady=2, sticky="we")
        self.vline_1.grid(row=grid_row, column=1, sticky="nsew")
        self.start_entry.grid(row=grid_row, column=2, padx=4, pady=2, sticky="we")
        self.vline_2.grid(row=grid_row, column=3, sticky="nsew")
        self.end_frame.grid(row=grid_row, column=4, padx=2, pady=2, sticky="nsew")
        self.vline_3.grid(row=grid_row, column=5, sticky="nsew")
        self.switch_btn.grid(row=grid_row, column=6, padx=2, pady=2)
        self.vline_4.grid(row=grid_row, column=7, sticky="nsew")
        self.count_frame.grid(row=grid_row, column=8, padx=2, pady=2, sticky="nsew")

        self.end_frame.columnconfigure(0, weight=1)
        self.count_frame.columnconfigure(0, weight=1)

        self.end_entry.grid(row=0, column=0, sticky="we")
        self.count_entry.grid(row=0, column=0, sticky="we")

        # Overlay X centered within the entry area
        self.end_block.place(relx=0.5, rely=0.5, anchor="center")
        self.count_block.place(relx=0.5, rely=0.5, anchor="center")

    def set_mode(self, use_end: bool):
        self.mode_use_end = use_end
        if use_end:
            self.end_entry.state(["!disabled"])
            self.count_entry.state(["disabled"])
            self.end_block.place_forget()
            self.count_block.place(relx=0.5, rely=0.5, anchor="center")
        else:
            self.end_entry.state(["disabled"])
            self.count_entry.state(["!disabled"])
            self.count_block.place_forget()
            self.end_block.place(relx=0.5, rely=0.5, anchor="center")

    def toggle_mode(self):
        if not self.confirm_switch(self.mode_use_end):
            return
        if self.mode_use_end:
            self.end_var.set("")
            self.set_mode(False)
        else:
            self.count_var.set("")
            self.set_mode(True)
        self.on_change()

    def destroy(self):
        self.filename_entry.destroy()
        self.start_entry.destroy()
        self.end_frame.destroy()
        self.switch_btn.destroy()
        self.count_frame.destroy()
        self.vline_1.destroy()
        self.vline_2.destroy()
        self.vline_3.destroy()
        self.vline_4.destroy()


def parse_int(value):
    if value is None or value == "":
        return None
    try:
        n = int(value)
    except ValueError:
        raise ValueError("not_int")
    return n


def build_ui():
    root = tk.Tk()
    root.title("PDF Cutter")
    root.geometry("860x520")

    main = ttk.Frame(root, padding=10)
    main.pack(fill="both", expand=True)

    lang = {"code": "ko"}
    texts = {
        "ko": {
            "title": "PDF Cutter",
            "input": "인풋 파일",
            "output": "아웃풋 경로",
            "browse": "찾기",
            "filename": "파일명",
            "start": "시작 페이지",
            "end": "종료 페이지",
            "switch": "전환",
            "count": "페이지 수",
            "add_row": "+ 행 추가",
            "offset": "페이지 오프셋 (PDF와 목차 페이지가 다를 때)",
            "append": "파일명에 페이지 범위 추가 (예: Filename_1p_to_5p.pdf)",
            "run": "실행",
            "toggle": "한/A",
            "err_title": "오류",
            "err_input_required": "인풋 파일을 선택하세요.",
            "err_input_not_found": "인풋 파일을 찾을 수 없습니다.",
            "err_output_required": "아웃풋 경로를 선택하세요.",
            "err_output_create": "아웃풋 폴더 생성 실패: {detail}",
            "err_offset_int": "페이지 오프셋은 정수여야 합니다.",
            "err_pdf_open": "PDF 열기 실패: {detail}",
            "err_row_int": "행 {row}: {field}은(는) 정수여야 합니다.",
            "err_row_start_required": "행 {row}: 시작 페이지는 필수입니다.",
            "err_row_start_min": "행 {row}: 시작 페이지는 1 이상이어야 합니다.",
            "err_row_end_mode": "행 {row}: End Page 모드에서는 페이지 수를 비워야 합니다.",
            "err_row_count_mode": "행 {row}: 페이지 수 모드에서는 End Page를 비워야 합니다.",
            "err_row_end_min": "행 {row}: 종료 페이지는 1 이상이어야 합니다.",
            "err_row_end_lt": "행 {row}: 종료 페이지는 시작 페이지보다 작을 수 없습니다.",
            "err_row_count_min": "행 {row}: 페이지 수는 1 이상이어야 합니다.",
            "err_row_required": "최소 1개 이상의 유효한 행이 필요합니다.",
            "err_start_oob": "행 {row}: 시작 페이지가 범위를 벗어났습니다.",
            "err_end_oob": "행 {row}: 종료 페이지가 범위를 벗어났습니다.",
            "err_end_before": "행 {row}: 보정 후 종료 페이지가 시작 페이지보다 이릅니다.",
            "done_title": "완료",
            "done_saved": "{count}개 파일을 저장했습니다:\n{path}",
        },
        "en": {
            "title": "PDF Cutter",
            "input": "Input File",
            "output": "Output Folder",
            "browse": "Browse",
            "filename": "Filename",
            "start": "Start Page",
            "end": "End Page",
            "switch": "Switch",
            "count": "Page Count",
            "add_row": "+ Add Row",
            "offset": "Page Offset (use when PDF page numbers differ)",
            "append": "Append page range to filename (e.g., Filename_1p_to_5p.pdf)",
            "run": "Run",
            "toggle": "A/한",
            "err_title": "Error",
            "err_input_required": "Input file is required.",
            "err_input_not_found": "Input file not found.",
            "err_output_required": "Output folder is required.",
            "err_output_create": "Cannot create output folder: {detail}",
            "err_offset_int": "Page Offset must be an integer.",
            "err_pdf_open": "Failed to open PDF: {detail}",
            "err_row_int": "Row {row}: {field} must be an integer.",
            "err_row_start_required": "Row {row}: Start Page is required.",
            "err_row_start_min": "Row {row}: Start Page must be >= 1.",
            "err_row_end_mode": "Row {row}: Page Count must be empty in End Page mode.",
            "err_row_count_mode": "Row {row}: End Page must be empty in Page Count mode.",
            "err_row_end_min": "Row {row}: End Page must be >= 1.",
            "err_row_end_lt": "Row {row}: End Page cannot be less than Start Page.",
            "err_row_count_min": "Row {row}: Page Count must be > 0.",
            "err_row_required": "At least one valid row is required.",
            "err_start_oob": "Row {row}: Start Page out of range.",
            "err_end_oob": "Row {row}: End Page out of range.",
            "err_end_before": "Row {row}: End Page earlier than Start Page after offset.",
            "done_title": "Done",
            "done_saved": "Saved {count} files to:\n{path}",
        },
    }

    def t(key: str) -> str:
        return texts[lang["code"]][key]

    def msg(key: str, **kwargs) -> str:
        template = texts[lang["code"]][key]
        return template.format(**kwargs)

    def apply_texts():
        root.title(t("title"))
        input_label_var.set(t("input"))
        output_label_var.set(t("output"))
        browse_in_var.set(t("browse"))
        browse_out_var.set(t("browse"))
        header_filename_var.set(t("filename"))
        header_start_var.set(t("start"))
        header_end_var.set(t("end"))
        header_switch_var.set(t("switch"))
        header_count_var.set(t("count"))
        add_row_var.set(t("add_row"))
        offset_label_var.set(t("offset"))
        append_label_var.set(t("append"))
        run_var.set(t("run"))
        # language buttons are handled separately

    def set_lang(code: str):
        lang["code"] = code
        apply_texts()
        if code == "ko":
            btn_ko.config(relief="sunken")
            btn_en.config(relief="raised")
        else:
            btn_ko.config(relief="raised")
            btn_en.config(relief="sunken")

    def confirm_switch(use_end_mode: bool) -> bool:
        if lang["code"] == "ko":
            title = "모드 전환"
            if use_end_mode:
                msg = "페이지 수 모드로 전환할까요?\nEnd Page 값은 비워집니다."
            else:
                msg = "End Page 모드로 전환할까요?\n페이지 수 값은 비워집니다."
        else:
            title = "Switch Mode"
            if use_end_mode:
                msg = "Switch to Page Count mode?\nEnd Page will be cleared."
            else:
                msg = "Switch to End Page mode?\nPage Count will be cleared."
        return messagebox.askyesno(title, msg)

    # Input/output
    io_frame = ttk.Frame(main)
    io_frame.pack(fill="x")

    input_var = tk.StringVar()
    output_var = tk.StringVar()

    input_label_var = tk.StringVar()
    output_label_var = tk.StringVar()
    browse_in_var = tk.StringVar()
    browse_out_var = tk.StringVar()

    ttk.Label(io_frame, textvariable=input_label_var).grid(row=0, column=0, sticky="w")
    input_entry = ttk.Entry(io_frame, textvariable=input_var, width=70)
    input_entry.grid(row=0, column=1, padx=6, pady=4, sticky="we")

    def pick_input():
        path = filedialog.askopenfilename(title="Select PDF", filetypes=[("PDF files", "*.pdf")])
        if path:
            input_var.set(path)
            base = os.path.splitext(os.path.basename(path))[0]
            input_base["value"] = base
            pattern = re.compile(r"(Filename|파일명)_\d+")
            for idx, row in enumerate(rows, start=1):
                current = row.filename_var.get().strip()
                if (not current) or pattern.fullmatch(current):
                    row.filename_var.set(f"{base}_{idx}")

    ttk.Button(io_frame, textvariable=browse_in_var, command=pick_input).grid(row=0, column=2, padx=4)

    ttk.Label(io_frame, textvariable=output_label_var).grid(row=1, column=0, sticky="w")
    output_entry = ttk.Entry(io_frame, textvariable=output_var, width=70)
    output_entry.grid(row=1, column=1, padx=6, pady=4, sticky="we")

    def pick_output():
        path = filedialog.askdirectory(title="Select output folder")
        if path:
            output_var.set(path)

    ttk.Button(io_frame, textvariable=browse_out_var, command=pick_output).grid(row=1, column=2, padx=4)

    io_frame.columnconfigure(1, weight=1)

    ttk.Separator(main, orient="horizontal").pack(fill="x", pady=8)

    # Table
    table_frame = tk.Frame(main, bd=1, relief="solid")
    table_frame.pack(fill="both", expand=True)

    table_frame.columnconfigure(0, weight=3)
    table_frame.columnconfigure(1, weight=0)
    table_frame.columnconfigure(2, weight=1, minsize=60, uniform="pagecols")
    table_frame.columnconfigure(3, weight=0)
    table_frame.columnconfigure(4, weight=1, minsize=60, uniform="pagecols")
    table_frame.columnconfigure(5, weight=0)
    table_frame.columnconfigure(6, weight=0, minsize=34)
    table_frame.columnconfigure(7, weight=0)
    table_frame.columnconfigure(8, weight=1, minsize=60, uniform="pagecols")

    header_filename_var = tk.StringVar()
    header_start_var = tk.StringVar()
    header_end_var = tk.StringVar()
    header_switch_var = tk.StringVar()
    header_count_var = tk.StringVar()

    tk.Label(table_frame, textvariable=header_filename_var, bg="#efefef", anchor="w").grid(row=0, column=0, padx=6, pady=4, sticky="we")
    tk.Frame(table_frame, bg="#d0d0d0", width=1).grid(row=0, column=1, sticky="nsew")
    tk.Label(table_frame, textvariable=header_start_var, bg="#efefef", anchor="center").grid(row=0, column=2, padx=6, pady=4, sticky="we")
    tk.Frame(table_frame, bg="#d0d0d0", width=1).grid(row=0, column=3, sticky="nsew")
    tk.Label(table_frame, textvariable=header_end_var, bg="#efefef", anchor="center").grid(row=0, column=4, padx=6, pady=4, sticky="we")
    tk.Frame(table_frame, bg="#d0d0d0", width=1).grid(row=0, column=5, sticky="nsew")
    tk.Label(table_frame, textvariable=header_switch_var, bg="#efefef", anchor="center").grid(row=0, column=6, padx=2, pady=4, sticky="we")
    tk.Frame(table_frame, bg="#d0d0d0", width=1).grid(row=0, column=7, sticky="nsew")
    tk.Label(table_frame, textvariable=header_count_var, bg="#efefef", anchor="center").grid(row=0, column=8, padx=6, pady=4, sticky="we")

    ttk.Separator(table_frame, orient="horizontal").grid(row=1, column=0, columnspan=9, sticky="we")

    rows = []
    input_base = {"value": ""}

    def on_row_change():
        pass

    def add_row():
        idx = len(rows)
        grid_row = 2 + idx
        default_name = None
        if input_base["value"]:
            default_name = f"{input_base['value']}_{idx + 1}"
        row = SplitRow(table_frame, idx, grid_row, on_row_change, confirm_switch, default_name=default_name)
        rows.append(row)
        row.set_mode(True)

    add_row()

    # Controls under table
    controls = ttk.Frame(main)
    controls.pack(fill="x", pady=6)

    add_row_var = tk.StringVar()
    add_btn = ttk.Button(controls, textvariable=add_row_var, command=add_row)
    add_btn.pack(fill="x")

    ttk.Separator(controls, orient="horizontal").pack(fill="x", pady=6)

    offset_row = ttk.Frame(controls)
    offset_row.pack(fill="x")
    offset_var = tk.StringVar(value="0")
    offset_label_var = tk.StringVar()
    ttk.Label(offset_row, textvariable=offset_label_var).pack(side="left")
    ttk.Entry(offset_row, textvariable=offset_var, width=8, justify="center").pack(side="left", padx=6)

    ttk.Separator(controls, orient="horizontal").pack(fill="x", pady=6)

    append_row = ttk.Frame(controls)
    append_row.pack(fill="x")
    append_var = tk.BooleanVar(value=True)
    append_label_var = tk.StringVar()
    ttk.Label(append_row, textvariable=append_label_var).pack(side="left")
    ttk.Checkbutton(append_row, text="", variable=append_var).pack(side="left", padx=(8, 0))

    def validate_rows(total_pages):
        specs = []
        errors = []
        for i, row in enumerate(rows, start=1):
            filename = row.filename_var.get().strip()
            start_s = row.start_var.get().strip()
            end_s = row.end_var.get().strip()
            count_s = row.count_var.get().strip()

            if not start_s and not end_s and not count_s:
                # Empty row, skip
                continue

            try:
                start = parse_int(start_s)
            except ValueError:
                errors.append(msg("err_row_int", row=i, field=t("start")))
                continue
            try:
                end = parse_int(end_s)
            except ValueError:
                errors.append(msg("err_row_int", row=i, field=t("end")))
                continue
            try:
                count = parse_int(count_s)
            except ValueError:
                errors.append(msg("err_row_int", row=i, field=t("count")))
                continue

            if start is None:
                errors.append(msg("err_row_start_required", row=i))
                continue
            if start <= 0:
                errors.append(msg("err_row_start_min", row=i))
                continue

            if row.mode_use_end:
                if count is not None:
                    errors.append(msg("err_row_end_mode", row=i))
                    continue
            else:
                if end is not None:
                    errors.append(msg("err_row_count_mode", row=i))
                    continue

            if end is not None:
                if end <= 0:
                    errors.append(msg("err_row_end_min", row=i))
                    continue
                if end < start:
                    errors.append(msg("err_row_end_lt", row=i))
                    continue
            if count is not None:
                if count <= 0:
                    errors.append(msg("err_row_count_min", row=i))
                    continue

            specs.append((filename, start, end, count, row.mode_use_end))

        if not specs:
            errors.append(msg("err_row_required"))

        return specs, errors

    def run_split():
        pdf_path = input_var.get().strip()
        out_dir = output_var.get().strip()

        if not pdf_path:
            messagebox.showerror(msg("err_title"), msg("err_input_required"))
            return
        if not os.path.isfile(pdf_path):
            messagebox.showerror(msg("err_title"), msg("err_input_not_found"))
            return
        if not out_dir:
            messagebox.showerror(msg("err_title"), msg("err_output_required"))
            return
        if not os.path.isdir(out_dir):
            try:
                os.makedirs(out_dir, exist_ok=True)
            except Exception as e:
                messagebox.showerror(msg("err_title"), msg("err_output_create", detail=e))
                return

        try:
            offset = int(offset_var.get().strip() or "0")
        except ValueError:
            messagebox.showerror(msg("err_title"), msg("err_offset_int"))
            return

        try:
            reader = PdfReader(pdf_path)
        except Exception as e:
            messagebox.showerror(msg("err_title"), msg("err_pdf_open", detail=e))
            return

        total_pages = len(reader.pages)
        specs, errors = validate_rows(total_pages)
        if errors:
            messagebox.showerror(msg("err_title"), "\n".join(errors))
            return

        # Fill missing end using next row's start - 1, or total pages for last row
        enriched = []
        for i, (filename, start, end, count, mode_use_end) in enumerate(specs):
            if end is None and count is None:
                if i < len(specs) - 1:
                    end = specs[i + 1][1] - 1
                else:
                    end = total_pages
            if end is None and count is not None:
                end = start + count - 1
            enriched.append((filename, start, end))

        saved = []
        for idx, (filename, start, end) in enumerate(enriched, start=1):

            start_idx = (start + offset) - 1
            end_idx = (end + offset) - 1

            if start_idx < 0 or start_idx >= total_pages:
                messagebox.showerror(msg("err_title"), msg("err_start_oob", row=idx))
                return
            if end_idx < start_idx:
                messagebox.showerror(msg("err_title"), msg("err_end_before", row=idx))
                return
            if end_idx >= total_pages:
                messagebox.showerror(msg("err_title"), msg("err_end_oob", row=idx))
                return

            writer = PdfWriter()
            for p in range(start_idx, end_idx + 1):
                writer.add_page(reader.pages[p])

            base_name = safe_filename(filename) if filename else f"Filename_{idx}"
            if append_var.get():
                base_name = f"{base_name}_{start}p_to_{end}p"

            out_path = os.path.join(out_dir, f"{base_name}.pdf")
            with open(out_path, "wb") as f:
                writer.write(f)

            saved.append(out_path)

        messagebox.showinfo(msg("done_title"), msg("done_saved", count=len(saved), path=out_dir))

    ttk.Separator(main, orient="horizontal").pack(fill="x", pady=8)

    run_frame = ttk.Frame(main)
    run_frame.pack(fill="x")
    run_var = tk.StringVar()
    ttk.Button(run_frame, textvariable=run_var, command=run_split).pack(side="right")

    # Language toggle buttons (bottom-left)
    lang_wrap = ttk.Frame(run_frame)
    lang_wrap.pack(side="left")
    btn_en = tk.Button(lang_wrap, text="A", width=3, relief="sunken", command=lambda: set_lang("en"))
    btn_ko = tk.Button(lang_wrap, text="한", width=3, relief="raised", command=lambda: set_lang("ko"))
    btn_en.pack(side="left")
    btn_ko.pack(side="left")

    apply_texts()
    set_lang("en")

    return root


def main():
    root = build_ui()
    root.mainloop()


if __name__ == "__main__":
    main()










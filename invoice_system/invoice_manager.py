#!/usr/bin/env python3
"""Invoice folder scanner that writes records into an Excel template."""

from __future__ import annotations

import argparse
import csv
import datetime as dt
import os
import re
import shutil
import subprocess
import sys
import tempfile
import unicodedata
from copy import copy
from dataclasses import dataclass
from pathlib import Path

from openpyxl import load_workbook


SUPPORTED_EXTS = {".pdf", ".png", ".jpg", ".jpeg", ".webp", ".bmp", ".tif", ".tiff"}
SUMMARY_EXTS = {".xlsx", ".xlsm"}
HEADER_ROW = 2
DATA_START_ROW = 3


def app_root() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent.parent


ROOT = app_root()


def default_template_path() -> Path:
    env_template = os.environ.get("INVOICE_TEMPLATE", "").strip()
    if env_template:
        return Path(env_template).expanduser()

    candidates = [
        ROOT / "发票模板.xlsx",
        ROOT / "windows_package" / "发票模板.xlsx",
    ]
    for candidate in candidates:
        if candidate.exists():
            return candidate
    return candidates[0]


DEFAULT_TEMPLATE = default_template_path()


def find_tool(name: str) -> str | None:
    executable = f"{name}.exe" if os.name == "nt" else name
    bundled_candidates = [
        ROOT / "tools" / "poppler" / "Library" / "bin" / executable,
        ROOT / "tools" / "poppler" / "bin" / executable,
        ROOT / "tools" / "tesseract" / executable,
        ROOT / "tools" / executable,
    ]
    for candidate in bundled_candidates:
        if candidate.exists():
            return str(candidate)
    return shutil.which(name)


@dataclass
class InvoiceRecord:
    source: Path
    reimburser: str = ""
    project: str = ""
    buyer_name: str = ""
    seller_name: str = ""
    amount: float | None = None
    invoice_date: str = ""
    invoice_no: str = ""
    remark: str = ""
    raw_text: str = ""

    @property
    def needs_review(self) -> bool:
        return (
            self.amount is None
            or not self.invoice_date
            or not self.invoice_no
            or not self.project
            or not self.buyer_name
        )


def run_command(args: list[str]) -> str:
    proc = subprocess.run(args, capture_output=True, text=True, check=False)
    if proc.returncode != 0:
        return ""
    return proc.stdout.strip()


def extract_text(path: Path) -> str:
    """Use available local text/OCR tools, then fall back to sidecar text files."""
    sidecar = path.with_suffix(".txt")
    if sidecar.exists():
        return sidecar.read_text(encoding="utf-8", errors="ignore")

    pdftotext = find_tool("pdftotext")
    pdftoppm = find_tool("pdftoppm")
    tesseract = find_tool("tesseract")

    if path.suffix.lower() == ".pdf" and pdftotext:
        text = run_command([pdftotext, "-layout", str(path), "-"])
        if text:
            return text

    if path.suffix.lower() == ".pdf" and pdftoppm and tesseract:
        text = ocr_pdf(path)
        if text:
            return text

    if path.suffix.lower() != ".pdf" and tesseract:
        lang = "chi_sim+eng"
        text = run_command([tesseract, str(path), "stdout", "-l", lang])
        if text:
            return text

    return ""


def ocr_pdf(path: Path) -> str:
    lang = "chi_sim+eng"
    pdftoppm = find_tool("pdftoppm")
    tesseract = find_tool("tesseract")
    if not pdftoppm or not tesseract:
        return ""
    with tempfile.TemporaryDirectory(prefix="invoice_ocr_") as tmp:
        prefix = str(Path(tmp) / "page")
        proc = subprocess.run(
            [pdftoppm, "-png", "-r", "300", str(path), prefix],
            capture_output=True,
            text=True,
            check=False,
        )
        if proc.returncode != 0:
            return ""

        texts = []
        for image in sorted(Path(tmp).glob("page-*.png")):
            text = run_command([tesseract, str(image), "stdout", "-l", lang])
            if text:
                texts.append(text)
        return "\n".join(texts).strip()


def normalize_text(text: str) -> str:
    # Some tax invoice PDFs encode ordinary Chinese characters as Unicode
    # radicals (for example: “⻓” for “长” and “⽹” for “网”).
    # NFKC fixes the Kangxi variants; the CJK radical supplement character
    # U+2ED3 needs an explicit mapping.
    text = unicodedata.normalize("NFKC", text).translate({0x2ED3: "长"})
    text = text.replace("\u3000", " ")
    text = re.sub(r"[ \t]+", " ", text)
    return text


def parse_amount(text: str, filename: str) -> float | None:
    haystack = normalize_text(text)
    patterns = [
        r"价税合计[\s\S]{0,120}[（(]小写[）)]?\s*[¥￥]?\s*([0-9]+(?:\.[0-9]{1,2})?)",
        r"[（(]小写[）)]?\s*[¥￥]?\s*([0-9]+(?:\.[0-9]{1,2})?)",
    ]
    for pattern in patterns:
        match = re.search(pattern, haystack)
        if match:
            return float(match.group(1))
    return None


def parse_date(text: str, filename: str) -> str:
    haystack = normalize_text(text)
    patterns = [
        r"开票日期[:：]?\s*(\d{4})[年./-](\d{1,2})[月./-](\d{1,2})日?",
    ]
    for pattern in patterns:
        match = re.search(pattern, haystack)
        if match:
            y, m, d = (int(x) for x in match.groups())
            return f"{y:04d}-{m:02d}-{d:02d}"
    return ""


def parse_invoice_no(text: str, filename: str) -> str:
    haystack = normalize_text(text)
    patterns = [
        r"发票号码[:：]?\s*([0-9 ]{8,40})",
        r"发票号[:：]?\s*([0-9 ]{8,40})",
    ]
    for pattern in patterns:
        match = re.search(pattern, haystack)
        if match:
            number = re.sub(r"\D", "", match.group(1))
            if len(number) >= 8:
                return number
    return ""


def clean_company_name(value: str) -> str:
    stop_words = [
        "纳税人识别号",
        "统一社会信用代码",
        "地址",
        "电话",
        "开户行",
        "账号",
        "销售方",
        "销方",
        "购买方",
        "购方",
        "项目名称",
    ]
    cleaned = value.strip()
    for word in stop_words:
        cleaned = cleaned.split(word, 1)[0].strip()
    cleaned = re.sub(r"^(?:名\s*称|购买方名称|购方名称)[:：]?", "", cleaned).strip()
    cleaned = re.sub(r"\s+", "", cleaned)
    return cleaned.strip("：:，,。；; ")


def parse_buyer_name(text: str, filename: str) -> str:
    haystack = normalize_text(text)
    patterns = [
        r"购买方信息[\s\S]{0,120}?名称[:：]?\s*([^\n\r]{2,80})",
        r"购买方[\s\S]{0,80}?名称[:：]?\s*([^\n\r]{2,80})",
        r"购\s*买\s*方\s*名\s*称[:：]?\s*([^\n\r]{2,80})",
        r"购\s*名\s*称[:：]?\s*([^\n\r]{2,80}?)(?:\s{2,}|销\s*名\s*称|$)",
        r"买\s*名\s*称[:：]?\s*([^\n\r]{2,80}?)(?:\s{2,}|售\s*名\s*称|$)",
        r"购方名称[:：]?\s*([^\n\r]{2,80})",
        r"付款方[:：]?\s*([^\n\r]{2,80})",
    ]
    for pattern in patterns:
        match = re.search(pattern, haystack)
        if not match:
            continue
        value = clean_company_name(match.group(1))
        if value:
            return value
    return ""


def parse_seller_name(text: str, filename: str) -> str:
    haystack = normalize_text(text)
    patterns = [
        r"销售方信息[\s\S]{0,120}?名称[:：]?\s*([^\n\r]{2,80})",
        r"销售方[\s\S]{0,80}?名称[:：]?\s*([^\n\r]{2,80})",
        r"销\s*售\s*方\s*名\s*称[:：]?\s*([^\n\r]{2,80})",
        r"销方名称[:：]?\s*([^\n\r]{2,80})",
        r"开票方[:：]?\s*([^\n\r]{2,80})",
        r"开票单位[:：]?\s*([^\n\r]{2,80})",
    ]
    for pattern in patterns:
        match = re.search(pattern, haystack)
        if not match:
            continue
        value = clean_company_name(match.group(1))
        if value:
            return value
    if "腾讯" in haystack:
        return "腾讯"
    return ""


def parse_phone_number(text: str, filename: str) -> str:
    combined = f"{filename} {normalize_text(text)}"
    match = re.search(r"(?<!\d)(1[3-9]\d{9})(?!\d)", combined)
    return match.group(1) if match else ""


def parse_project(text: str, filename: str) -> str:
    haystack = normalize_text(text)
    combined = f"{filename} {haystack}"
    if "微信认证" in combined or ("腾讯" in combined and "认证" in combined):
        return "微信认证费"

    keywords = [
        ("电话", "电话费"),
        ("话费", "电话费"),
        ("通信", "通讯费"),
        ("通讯", "通讯费"),
        ("餐饮", "餐饮费"),
        ("住宿", "住宿费"),
        ("交通", "交通费"),
        ("汽油", "交通费"),
        ("打车", "交通费"),
        ("快递", "快递费"),
        ("办公", "办公费"),
    ]
    for key, project in keywords:
        if key in combined:
            return project

    match = re.search(r"项目名称[:：]?\s*([^\n\r]{2,40})", haystack)
    if match:
        return match.group(1).strip()
    return ""


def build_remark(record: InvoiceRecord, phone_number: str) -> str:
    remarks = []
    if record.project == "微信认证费":
        account_name = record.source.stem.strip()
        if account_name:
            remarks.append(f"认证账号：{account_name}")

    if phone_number and record.project in {"电话费", "通讯费"}:
        remarks.append(f"号码：{phone_number}")

    missing = []
    if not record.project:
        missing.append("项目")
    if record.amount is None:
        missing.append("金额")
    if not record.invoice_date:
        missing.append("开票日期")
    if not record.invoice_no:
        missing.append("发票号码")
    if missing:
        remarks.append("待核对：" + "、".join(missing))
    return "；".join(remarks)


def build_record(path: Path, reimburser: str) -> InvoiceRecord:
    text = extract_text(path)
    filename = path.name
    phone_number = parse_phone_number(text, filename)
    record = InvoiceRecord(
        source=path,
        reimburser=reimburser,
        project=parse_project(text, filename),
        buyer_name=parse_buyer_name(text, filename),
        seller_name=parse_seller_name(text, filename),
        amount=parse_amount(text, filename),
        invoice_date=parse_date(text, filename),
        invoice_no=parse_invoice_no(text, filename),
        raw_text=text,
    )
    record.remark = build_remark(record, phone_number)
    return record


def normalize_invoice_no(value: object) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    return re.sub(r"\D", "", str(value))


def normalize_date_value(value: object) -> str:
    if value is None:
        return ""
    if isinstance(value, dt.datetime | dt.date):
        return value.strftime("%Y-%m-%d")
    text = str(value).strip()
    match = re.search(r"(\d{4})[年./-](\d{1,2})[月./-](\d{1,2})", text)
    if match:
        y, m, d = (int(x) for x in match.groups())
        return f"{y:04d}-{m:02d}-{d:02d}"
    return text


def normalize_amount_value(value: object) -> float | None:
    if value is None or value == "":
        return None
    if isinstance(value, int | float):
        return float(value)
    match = re.search(r"-?\d+(?:\.\d+)?", str(value).replace(",", ""))
    return float(match.group(0)) if match else None


def is_summary_row(ws, row: int) -> bool:
    values = [ws.cell(row, col).value for col in range(1, ws.max_column + 1)]
    return any(isinstance(value, str) and "汇总" in value for value in values)


def detect_header_map(ws) -> tuple[int, dict[str, int]] | None:
    aliases = {
        "reimburser": ("报销人", "报销人员", "申请人", "员工", "姓名"),
        "project": ("项目", "费用项目", "费用类型", "类别"),
        "amount": ("金额", "价税合计", "报销金额", "费用金额"),
        "invoice_date": ("开票日期", "发票日期", "日期"),
        "invoice_no": ("发票", "发票号码", "发票号", "票据号码"),
        "remark": ("备注", "说明", "认证账号", "用途"),
    }
    for row in range(1, min(ws.max_row, 20) + 1):
        mapping: dict[str, int] = {}
        for col in range(1, ws.max_column + 1):
            value = ws.cell(row, col).value
            if value is None:
                continue
            label = re.sub(r"\s+", "", str(value))
            for key, names in aliases.items():
                if key not in mapping and any(name in label for name in names):
                    mapping[key] = col
        if "invoice_no" in mapping:
            return row, mapping
    return None


def invoice_no_from_path(path: Path) -> str:
    match = re.search(r"(?<!\d)(\d{8,40})(?!\d)", path.stem)
    if match:
        return match.group(1)
    return normalize_invoice_no(path.stem)


def company_name_from_pdf_path(input_dir: Path, path: Path) -> str:
    rel = path.relative_to(input_dir)
    parent_parts = rel.parent.parts
    if not parent_parts:
        return ""
    if parent_parts[-1] in {"发票PDF", "PDF", "pdf", "发票"} and len(parent_parts) >= 2:
        return parent_parts[-2]
    return parent_parts[-1]


def build_pdf_index(input_dir: Path) -> dict[str, list[Path]]:
    index: dict[str, list[Path]] = {}
    for path in sorted(input_dir.rglob("*.pdf")):
        invoice_no = invoice_no_from_path(path)
        if not invoice_no:
            continue
        index.setdefault(invoice_no, []).append(path)
    return index


def append_missing_review(remarks: list[str], record: InvoiceRecord) -> None:
    missing = []
    if not record.project:
        missing.append("项目")
    if record.amount is None:
        missing.append("金额")
    if not record.invoice_date:
        missing.append("开票日期")
    if not record.invoice_no:
        missing.append("发票号码")
    if not record.buyer_name:
        missing.append("购买方公司")
    if missing:
        remarks.append("待核对：" + "、".join(missing))


def build_record_from_summary_row(
    ws,
    row: int,
    columns: dict[str, int],
    invoice_no: str,
    source: Path,
    input_dir: Path,
    pdf_record_cache: dict[Path, InvoiceRecord],
    match_count: int,
) -> InvoiceRecord:
    def cell_value(key: str) -> object:
        col = columns.get(key)
        return ws.cell(row, col).value if col else None

    if source.suffix.lower() == ".pdf":
        pdf_record = pdf_record_cache.get(source)
        if pdf_record is None:
            pdf_record = build_record(source, "")
            pdf_record_cache[source] = pdf_record
        record = InvoiceRecord(
            source=source,
            reimburser=pdf_record.reimburser,
            project=pdf_record.project,
            buyer_name=pdf_record.buyer_name,
            seller_name=pdf_record.seller_name,
            amount=pdf_record.amount,
            invoice_date=pdf_record.invoice_date,
            invoice_no=invoice_no or pdf_record.invoice_no,
            remark="",
            raw_text=pdf_record.raw_text,
        )
    else:
        record = InvoiceRecord(source=source, invoice_no=invoice_no)

    record.reimburser = str(cell_value("reimburser") or record.reimburser or "").strip()
    record.project = str(cell_value("project") or record.project or "").strip()
    amount = normalize_amount_value(cell_value("amount"))
    if amount is not None:
        record.amount = amount
    date_value = normalize_date_value(cell_value("invoice_date"))
    if date_value:
        record.invoice_date = date_value

    remarks = []
    remark_col = columns.get("remark")
    if remark_col:
        value = ws.cell(row, remark_col).value
        if value not in (None, ""):
            remarks.append(str(value).strip())

    if source.suffix.lower() != ".pdf":
        remarks.append("待核对：未找到对应PDF")
    elif not record.buyer_name:
        folder_company = company_name_from_pdf_path(input_dir, source)
        if folder_company:
            record.buyer_name = folder_company
            remarks.append("待核对：购买方公司来自PDF所在文件夹")
    if match_count > 1:
        remarks.append("待核对：发票号码匹配到多个PDF")
    append_missing_review(remarks, record)
    record.remark = "；".join(remark for remark in remarks if remark)
    return record


def read_records_from_summary_workbook(
    path: Path,
    input_dir: Path,
    pdf_index: dict[str, list[Path]],
    pdf_record_cache: dict[Path, InvoiceRecord],
) -> list[InvoiceRecord]:
    wb = load_workbook(path, data_only=True)
    records: list[InvoiceRecord] = []
    for ws in wb.worksheets:
        detected = detect_header_map(ws)
        if not detected:
            continue
        header_row, columns = detected
        for row in range(header_row + 1, ws.max_row + 1):
            if is_summary_row(ws, row):
                continue
            invoice_no = normalize_invoice_no(ws.cell(row, columns["invoice_no"]).value)
            if not invoice_no:
                continue
            matches = pdf_index.get(invoice_no, [])
            records.append(
                build_record_from_summary_row(
                    ws,
                    row,
                    columns,
                    invoice_no,
                    matches[0] if matches else path,
                    input_dir,
                    pdf_record_cache,
                    len(matches),
                )
            )
    return records


def scan_summary_workbooks(input_dir: Path) -> list[InvoiceRecord]:
    workbooks = sorted(
        p for p in input_dir.rglob("*") if p.is_file() and p.suffix.lower() in SUMMARY_EXTS and not p.name.startswith("~$")
    )
    if not workbooks:
        return []
    pdf_index = build_pdf_index(input_dir)
    pdf_record_cache: dict[Path, InvoiceRecord] = {}
    records: list[InvoiceRecord] = []
    for workbook in workbooks:
        records.extend(read_records_from_summary_workbook(workbook, input_dir, pdf_index, pdf_record_cache))
    return dedupe_records(records)


def dedupe_records(records: list[InvoiceRecord]) -> list[InvoiceRecord]:
    deduped = []
    seen = set()
    for record in records:
        key = (
            record.invoice_no,
            record.reimburser,
            record.project,
            record.amount,
            record.invoice_date,
            record.buyer_name,
            record.remark,
        )
        if key in seen:
            continue
        seen.add(key)
        deduped.append(record)
    return deduped


def copy_row_style(ws, source_row: int, target_row: int) -> None:
    for col in range(1, ws.max_column + 1):
        src = ws.cell(source_row, col)
        dst = ws.cell(target_row, col)
        if src.has_style:
            dst.font = copy(src.font)
            dst.fill = copy(src.fill)
            dst.border = copy(src.border)
            dst.alignment = copy(src.alignment)
            dst.number_format = src.number_format
            dst.protection = copy(src.protection)


def find_summary_row(ws) -> int | None:
    for row in range(DATA_START_ROW, ws.max_row + 1):
        values = [ws.cell(row, col).value for col in range(1, ws.max_column + 1)]
        if any(isinstance(value, str) and "汇总" in value for value in values):
            return row
        if any(isinstance(value, str) and value.startswith("=") for value in values):
            return row
    return None


def clear_template_data(ws) -> None:
    summary_row = find_summary_row(ws)
    end_row = (summary_row - 1) if summary_row else ws.max_row
    for row in range(DATA_START_ROW, end_row + 1):
        for col in range(2, ws.max_column + 1):
            ws.cell(row, col).value = None
            ws.cell(row, col).hyperlink = None
            ws.cell(row, col).comment = None


def refresh_summary_formula(ws, summary_row: int, last_data_row: int) -> None:
    if last_data_row < DATA_START_ROW:
        last_data_row = DATA_START_ROW
    for col in range(1, ws.max_column + 1):
        cell = ws.cell(summary_row, col)
        value = cell.value
        if isinstance(value, str) and value.startswith("=") and "SUM(" in value.upper():
            cell.value = f"=SUM(D{DATA_START_ROW}:D{last_data_row})"


def write_workbook(template: Path, output: Path, records: list[InvoiceRecord]) -> None:
    records = sort_records_for_summary(records)
    wb = load_workbook(template)
    ws = wb.active
    clear_template_data(ws)

    summary_row = find_summary_row(ws)
    if summary_row and DATA_START_ROW + len(records) > summary_row:
        rows_needed = DATA_START_ROW + len(records) - summary_row
        ws.insert_rows(summary_row, rows_needed)
        summary_row += rows_needed

    for idx, record in enumerate(records, start=1):
        row = DATA_START_ROW + idx - 1
        copy_row_style(ws, DATA_START_ROW, row)
        ws.cell(row, 1).value = idx
        ws.cell(row, 2).value = record.reimburser
        ws.cell(row, 3).value = record.project
        ws.cell(row, 4).value = record.amount
        ws.cell(row, 5).value = record.invoice_date
        ws.cell(row, 6).value = record.invoice_no
        ws.cell(row, 6).hyperlink = None
        ws.cell(row, 7).value = record.remark

    summary_row = find_summary_row(ws)
    if summary_row:
        refresh_summary_formula(ws, summary_row, DATA_START_ROW + len(records) - 1)

    output.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output)


def write_review_csv(output: Path, records: list[InvoiceRecord]) -> None:
    records = sort_records_for_summary(records)
    review_path = output.with_suffix(".待核对.csv")
    with review_path.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(["文件", "报销人", "项目", "购买方公司", "金额", "开票日期", "发票号码", "备注"])
        for record in records:
            writer.writerow(
                [
                    str(record.source),
                    record.reimburser,
                    record.project,
                    record.buyer_name,
                    "" if record.amount is None else record.amount,
                    record.invoice_date,
                    record.invoice_no,
                    record.remark,
                ]
            )


def buyer_group_name(record: InvoiceRecord) -> str:
    return record.buyer_name or "购买方公司待核对"


def safe_pdf_stem(value: str, fallback: str) -> str:
    cleaned = "".join(ch for ch in value.strip() if ch not in '/\\:*?"<>|')
    return cleaned or fallback


def unique_path(path: Path) -> Path:
    if not path.exists():
        return path
    counter = 2
    while True:
        candidate = path.with_name(f"{path.stem}_{counter}{path.suffix}")
        if not candidate.exists():
            return candidate
        counter += 1


def export_renamed_pdfs(records: list[InvoiceRecord], export_dir: Path) -> list[Path]:
    """Copy source PDFs into export_dir using invoice numbers as filenames."""
    records = sort_records_for_summary(records)
    export_dir.mkdir(parents=True, exist_ok=True)
    exported = []
    for idx, record in enumerate(records, start=1):
        if record.source.suffix.lower() != ".pdf":
            continue
        fallback = f"待核对_{idx}_{record.source.stem}"
        filename = safe_pdf_stem(record.invoice_no, fallback) + ".pdf"
        target = unique_path(export_dir / filename)
        shutil.copy2(record.source, target)
        exported.append(target)
    return exported


def group_records_by_buyer(records: list[InvoiceRecord]) -> dict[str, list[InvoiceRecord]]:
    groups: dict[str, list[InvoiceRecord]] = {}
    for record in sort_records_for_summary(records):
        groups.setdefault(buyer_group_name(record), []).append(record)
    return groups


def export_buyer_results(template: Path, output_dir: Path, month: str, records: list[InvoiceRecord]) -> dict:
    if output_dir.exists():
        shutil.rmtree(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    groups = group_records_by_buyer(records)
    company_dirs = []
    exported_pdf_count = 0

    for buyer_name, buyer_records in groups.items():
        buyer_dir = output_dir / safe_pdf_stem(buyer_name, "购买方公司待核对")
        buyer_dir.mkdir(parents=True, exist_ok=True)
        output = buyer_dir / f"{buyer_name}{month}汇总.xlsx"
        write_workbook(template, output, buyer_records)
        write_review_csv(output, buyer_records)
        pdf_dir = buyer_dir / "发票PDF"
        exported_pdf_count += len(export_renamed_pdfs(buyer_records, pdf_dir))
        company_dirs.append(buyer_dir)

    return {
        "company_count": len(groups),
        "pdf_count": exported_pdf_count,
        "company_dirs": company_dirs,
    }


def sort_records_for_summary(records: list[InvoiceRecord]) -> list[InvoiceRecord]:
    def key(record: InvoiceRecord) -> tuple[str, str, str, str, str]:
        buyer = buyer_group_name(record)
        return (
            buyer,
            record.reimburser,
            record.project,
            record.invoice_date,
            record.invoice_no or record.source.name,
        )

    return sorted(records, key=key)


def scan_folder(input_dir: Path, reimburser: str) -> list[InvoiceRecord]:
    summary_records = scan_summary_workbooks(input_dir)
    summary_invoice_nos = {record.invoice_no for record in summary_records if record.invoice_no}

    files = sorted(
        p for p in input_dir.rglob("*") if p.is_file() and p.suffix.lower() in SUPPORTED_EXTS
    )
    records = list(summary_records)
    for path in files:
        if summary_invoice_nos and invoice_no_from_path(path) in summary_invoice_nos:
            continue
        rel_path = path.relative_to(input_dir)
        file_reimburser = reimburser
        if len(rel_path.parts) > 1:
            file_reimburser = rel_path.parts[0]
        records.append(build_record(path, file_reimburser))
    return records


def default_output(input_dir: Path) -> Path:
    month_name = input_dir.name.strip() or dt.date.today().strftime("%Y-%m")
    return input_dir.parent / f"{month_name}月度发票管理表.xlsx"


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="自动整理发票并写入 Excel 模板")
    parser.add_argument("input_dir", help="存放 PDF/图片发票的月份文件夹")
    parser.add_argument("--template", default=str(DEFAULT_TEMPLATE), help="Excel 模板路径")
    parser.add_argument("--output", default="", help="输出 Excel 路径，默认放在输入文件夹旁边")
    parser.add_argument("--reimburser", default="", help="报销人姓名")
    return parser.parse_args(argv)


def main(argv: list[str]) -> int:
    args = parse_args(argv)
    input_dir = Path(args.input_dir).expanduser().resolve()
    template = Path(args.template).expanduser().resolve()
    output = Path(args.output).expanduser().resolve() if args.output else default_output(input_dir)

    if not input_dir.exists():
        print(f"输入文件夹不存在：{input_dir}", file=sys.stderr)
        return 2
    if not template.exists():
        print(f"模板不存在：{template}", file=sys.stderr)
        return 2

    records = scan_folder(input_dir, args.reimburser)
    if not records:
        print(f"没有找到支持的发票文件：{input_dir}")
        return 0

    write_workbook(template, output, records)
    write_review_csv(output, records)
    pdf_export_dir = output.parent / f"{output.stem}_按发票号码命名PDF"
    exported_pdfs = export_renamed_pdfs(records, pdf_export_dir)

    reviewed = sum(1 for r in records if r.needs_review)
    print(f"已生成：{output}")
    print(f"已导出重命名 PDF：{pdf_export_dir}（{len(exported_pdfs)} 个）")
    print(f"共处理 {len(records)} 个文件，其中 {reviewed} 个需要人工核对。")
    print(f"待核对清单：{output.with_suffix('.待核对.csv')}")
    if not find_tool("pdftotext") and not find_tool("tesseract"):
        print("提示：当前未检测到 pdftotext/tesseract，扫描型 PDF 和图片需要安装 OCR 工具后才能自动识别正文。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))

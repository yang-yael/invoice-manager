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
HEADER_ROW = 2
DATA_START_ROW = 3


def app_root() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent.parent


ROOT = app_root()
DEFAULT_TEMPLATE = ROOT / "发票模板.xlsx"


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
        return self.amount is None or not self.invoice_date or not self.invoice_no or not self.project


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
    files = sorted(
        p for p in input_dir.rglob("*") if p.is_file() and p.suffix.lower() in SUPPORTED_EXTS
    )
    records = []
    for path in files:
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

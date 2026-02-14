import csv
import glob
import os
import unicodedata
from datetime import date, datetime, time, timedelta
from copy import copy

from openpyxl import Workbook, load_workbook
from openpyxl.utils.datetime import from_excel
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


HEADER_ROW = 3
DATA_START_ROW = 4
OUTPUT_COLUMNS = [1, 3, 7, 8, 9, 11]


def _find_input_workbook(root_dir: str) -> str:
    target = os.path.join(root_dir, "Pasta de Viagens Itu.xlsx")
    if os.path.exists(target):
        return target
    candidates = [
        path
        for path in glob.glob(os.path.join(root_dir, "Pasta*.xlsx"))
        if not os.path.basename(path).startswith("~")
    ]
    if not candidates:
        raise FileNotFoundError(
            "Nao encontrei arquivo 'Pasta de Viagens Itu.xlsx' nem outro 'Pasta*.xlsx'."
        )
    candidates.sort(key=lambda p: os.path.getmtime(p), reverse=True)
    return candidates[0]


def _target_date(now: datetime) -> date:
    if now.time() >= time(22, 0):
        return (now + timedelta(days=1)).date()
    return now.date()


def _normalize_text(value: object) -> str:
    if value is None:
        return ""
    text = str(value).strip().upper()
    text = unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode("ascii")
    return " ".join(text.split())


def _normalize_header(value: object) -> str:
    text = _normalize_text(value)
    return text.replace("_", " ")


def _find_header_index(headers: list[object], name: str) -> int:
    target = _normalize_header(name)
    for idx, header in enumerate(headers):
        if _normalize_header(header) == target:
            return idx
    raise ValueError(f"Nao encontrei a coluna '{name}' no cabecalho.")


def _to_date(value, epoch) -> date | None:
    if value is None or value == "":
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, (int, float)):
        try:
            return from_excel(value, epoch).date()
        except Exception:
            return None
    if isinstance(value, str):
        text = value.strip()
        if not text:
            return None
        formats = (
            "%d/%m/%Y",
            "%d/%m/%Y %H:%M",
            "%d/%m/%Y %H:%M:%S",
            "%d-%m-%Y",
            "%d-%m-%Y %H:%M",
            "%d-%m-%Y %H:%M:%S",
            "%Y-%m-%d",
            "%Y-%m-%d %H:%M",
            "%Y-%m-%d %H:%M:%S",
            "%d.%m.%Y",
            "%d.%m.%Y %H:%M",
            "%d.%m.%Y %H:%M:%S",
            "%m/%d/%Y",
            "%m/%d/%Y %H:%M",
            "%m/%d/%Y %H:%M:%S",
            "%Y/%m/%d",
            "%Y/%m/%d %H:%M",
            "%Y/%m/%d %H:%M:%S",
        )
        for fmt in formats:
            try:
                return datetime.strptime(text, fmt).date()
            except ValueError:
                continue
    return None


def _format_csv_value(value) -> str:
    if value is None:
        return ""
    if isinstance(value, datetime):
        if value.time() == time(0, 0):
            return value.strftime("%d/%m/%Y")
        return value.strftime("%d/%m/%Y %H:%M")
    if isinstance(value, date):
        return value.strftime("%d/%m/%Y")
    if isinstance(value, time):
        return value.strftime("%H:%M")
    return str(value)


def main() -> None:
    root_dir = os.path.dirname(os.path.abspath(__file__))
    input_path = _find_input_workbook(root_dir)

    wb = load_workbook(input_path, data_only=True)
    ws = wb.worksheets[0]
    wb_fmt = load_workbook(input_path)
    ws_fmt = wb_fmt.worksheets[0]

    if ws.auto_filter:
        ws.auto_filter.ref = None
        for row_idx in range(1, (ws.max_row or 0) + 1):
            ws.row_dimensions[row_idx].hidden = False

    all_headers = [
        ws.cell(row=HEADER_ROW, column=col).value for col in range(1, 35)
    ]

    origin_idx = _find_header_index(all_headers, "ORIGEM")
    date_idx = _find_header_index(all_headers, "DATA_SAIDA")

    output_headers = [all_headers[col - 1] for col in OUTPUT_COLUMNS]
    output_headers = [
        "HORA SUB PICKING" if _normalize_header(h) == "HORA SAIDA" else h
        for h in output_headers
    ]

    target = _target_date(datetime.now())

    filtered_rows: list[list[object]] = []
    max_row = ws.max_row or 0
    for row_idx in range(DATA_START_ROW, max_row + 1):
        row_all = [
            ws.cell(row=row_idx, column=col).value for col in range(1, 35)
        ]

        factory_value = _normalize_text(row_all[origin_idx])
        if factory_value != "FABRICA ITU":
            continue

        date_value = _to_date(row_all[date_idx], wb.epoch)
        if date_value != target:
            continue

        row_filtered = [row_all[col - 1] for col in OUTPUT_COLUMNS]
        filtered_rows.append((row_idx, row_filtered))

    date_str = target.strftime("%d-%m-%Y")
    base_name = f"solicitacao de pickings {date_str}"
    excel_path = os.path.join(root_dir, f"{base_name}.xlsx")
    csv_path = os.path.join(root_dir, f"{base_name}.csv")

    if os.path.exists(excel_path):
        os.remove(excel_path)
    if os.path.exists(csv_path):
        os.remove(csv_path)

    out_wb = Workbook()
    out_ws = out_wb.active
    out_ws.title = "Solicitacao"

    num_cols = len(OUTPUT_COLUMNS)

    out_ws.append(["CDV E CABREUVA"])
    out_ws.merge_cells(f"A1:{get_column_letter(num_cols)}1")
    
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    title_cell = out_ws.cell(row=1, column=1)
    title_cell.font = Font(bold=True, color="FFFFFF", size=12)
    title_cell.fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    title_cell.border = thin_border

    out_ws.append(output_headers)
    for out_col_idx, src_col in enumerate(OUTPUT_COLUMNS, 1):
        src_cell = ws_fmt.cell(row=HEADER_ROW, column=src_col)
        dst_cell = out_ws.cell(row=2, column=out_col_idx)
        dst_cell.font = Font(bold=True, color="FFFFFF", size=10)
        dst_cell.fill = PatternFill(start_color="4A90E2", end_color="4A90E2", fill_type="solid")
        dst_cell.alignment = Alignment(horizontal="center", vertical="center")
        dst_cell.border = thin_border
        dst_cell.value = output_headers[out_col_idx - 1]

    for row_idx, row_filtered in filtered_rows:
        out_ws.append(row_filtered)
        out_row = out_ws.max_row
        for out_col_idx, src_col in enumerate(OUTPUT_COLUMNS, 1):
            src_cell = ws_fmt.cell(row=row_idx, column=src_col)
            dst_cell = out_ws.cell(row=out_row, column=out_col_idx)
            src_font = copy(src_cell.font) if src_cell.has_style and src_cell.font else Font()
            dst_cell.font = Font(size=8, color=src_font.color if src_font else "000000")
            dst_cell.number_format = copy(src_cell.number_format) if src_cell.has_style else None
            dst_cell.border = thin_border
            dst_cell.alignment = copy(src_cell.alignment) if src_cell.has_style and src_cell.alignment else None

    out_ws.row_dimensions[1].height = 20
    out_ws.row_dimensions[2].height = 18
    for row_idx in range(3, out_ws.max_row + 1):
        out_ws.row_dimensions[row_idx].height = 15

    col_widths = {}
    for col_idx in range(1, num_cols + 1):
        max_length = 0
        col_letter = get_column_letter(col_idx)
        for row_idx in range(1, out_ws.max_row + 1):
            cell = out_ws.cell(row=row_idx, column=col_idx)
            if cell.value:
                cell_length = len(str(cell.value))
                if cell_length > max_length:
                    max_length = cell_length
        adjusted_width = min(max_length + 2, 40)
        col_widths[col_letter] = adjusted_width
        out_ws.column_dimensions[col_letter].width = adjusted_width

    out_wb.save(excel_path)

    with open(csv_path, "w", newline="", encoding="utf-8-sig") as csv_file:
        writer = csv.writer(csv_file, delimiter=";")
        writer.writerow(output_headers)
        for _, row_filtered in filtered_rows:
            writer.writerow([_format_csv_value(value) for value in row_filtered])

    print(f"Entrada: {os.path.basename(input_path)}")
    print(f"Data alvo: {target.strftime('%d/%m/%Y')}")
    print(f"Linhas filtradas: {len(filtered_rows)}")
    print(f"Excel: {excel_path}")
    print(f"CSV: {csv_path}")


if __name__ == "__main__":
    main()

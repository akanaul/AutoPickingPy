import csv
import glob
import logging
import os
import shutil
import subprocess
import sys
import unicodedata
import json
from datetime import date, datetime, time, timedelta
from pathlib import Path

from openpyxl import load_workbook, Workbook
from openpyxl.utils.datetime import from_excel
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment, Border, Side, Font, PatternFill

# Import license manager
try:
    from license_manager import check_license_and_authorize
except ImportError:
    # Fallback if license_manager not available
    def check_license_and_authorize():
        print("WARNING: License manager not available. Running without license check.")
        return True

try:
    from odf.opendocument import OpenDocumentSpreadsheet
    from odf.table import Table, TableRow, TableCell, TableColumn, CoveredTableCell
    from odf.text import P
    from odf.style import (
        Style,
        TextProperties,
        ParagraphProperties,
        TableCellProperties,
        TableColumnProperties,
    )
    from odf.number import DateStyle, TimeStyle, Day, Month, Year, Hours, Minutes, Text
except Exception:  # pragma: no cover - dependencia opcional
    OpenDocumentSpreadsheet = None


HEADER_ROW = 3
DATA_START_ROW = 4
OUTPUT_COLUMNS = [1, 3, 7, 8, 9, 11]
OUTPUT_HEADERS = [
    "DT",
    "STO",
    "DESTINO",
    "DATA SAIDA",
    "HORA SUB PICKING",
    "HORA SAIDA CARGA",
]



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


def _load_output_template(root_dir: str) -> tuple[Workbook, object]:
    wb = Workbook()
    ws = wb.active
    ws.title = "CABREUVA"
    
    ws.row_dimensions[1].height = 25
    ws.row_dimensions[2].height = 20
    
    return wb, ws


def _find_last_data_row(ws, max_col: int) -> int:
    max_row = ws.max_row or 1
    for row_idx in range(max_row, 0, -1):
        for col_idx in range(1, max_col + 1):
            value = ws.cell(row=row_idx, column=col_idx).value
            if value not in (None, ""):
                return row_idx
    return 1


def _write_ods(path: str, headers: list[str], rows: list[list[object]]) -> None:
    if OpenDocumentSpreadsheet is None:
        logging.warning("odfpy nao instalado; pulando geracao do ODS")
        return

    def _ods_date_value(value: date) -> str:
        return value.strftime("%Y-%m-%d")

    def _ods_time_value(value: time) -> str:
        return f"PT{value.hour:02d}H{value.minute:02d}M{value.second:02d}S"

    doc = OpenDocumentSpreadsheet()

    title_style = Style(name="TitleCell", family="table-cell")
    title_style.addElement(
        TableCellProperties(backgroundcolor="#2E5C8A", border="0.5pt solid #000000")
    )
    title_style.addElement(
        ParagraphProperties(textalign="center", verticalalign="middle")
    )
    title_style.addElement(
        TextProperties(color="#FFFFFF", fontweight="bold", fontsize="12pt")
    )

    header_style = Style(name="HeaderCell", family="table-cell")
    header_style.addElement(
        TableCellProperties(backgroundcolor="#4A90E2", border="0.5pt solid #000000")
    )
    header_style.addElement(
        ParagraphProperties(textalign="center", verticalalign="middle")
    )
    header_style.addElement(
        TextProperties(color="#FFFFFF", fontweight="bold", fontsize="10pt")
    )

    data_style = Style(name="DataCell", family="table-cell")
    data_style.addElement(
        TableCellProperties(border="0.5pt solid #000000")
    )
    data_style.addElement(
        ParagraphProperties(textalign="center", verticalalign="middle")
    )
    data_style.addElement(TextProperties(fontsize="8pt"))

    date_number = DateStyle(name="DateStyle")
    date_number.addElement(Day(style="long"))
    date_number.addElement(Text(text="/"))
    date_number.addElement(Month(style="long"))
    date_number.addElement(Text(text="/"))
    date_number.addElement(Year(style="long"))

    time_number = TimeStyle(name="TimeStyle")
    time_number.addElement(Hours(style="long"))
    time_number.addElement(Text(text=":"))
    time_number.addElement(Minutes(style="long"))

    date_cell_style = Style(
        name="DateCell", family="table-cell", parentstylename="DataCell", datastylename="DateStyle"
    )
    time_cell_style = Style(
        name="TimeCell", family="table-cell", parentstylename="DataCell", datastylename="TimeStyle"
    )

    doc.styles.addElement(title_style)
    doc.styles.addElement(header_style)
    doc.styles.addElement(data_style)
    doc.styles.addElement(date_number)
    doc.styles.addElement(time_number)
    doc.styles.addElement(date_cell_style)
    doc.styles.addElement(time_cell_style)

    max_lengths = [len(str(h)) for h in headers]
    for row in rows:
        for idx, value in enumerate(row):
            length = len(_format_csv_value(value))
            if length > max_lengths[idx]:
                max_lengths[idx] = length

    table = Table(name="Solicitacao")
    for width in max_lengths:
        width_cm = min(12.0, max(2.5, width * 0.28 + 0.8))
        col_style = Style(name=f"Col{width}", family="table-column")
        col_style.addElement(TableColumnProperties(columnwidth=f"{width_cm:.2f}cm"))
        doc.automaticstyles.addElement(col_style)
        table.addElement(TableColumn(stylename=col_style))

    title_row = TableRow()
    title_cell = TableCell(stylename=title_style, numbercolumnsspanned=len(headers) - 1)
    title_cell.addElement(P(text="CDV E CABREUVA"))
    title_row.addElement(title_cell)
    for _ in range(len(headers) - 2):
        title_row.addElement(CoveredTableCell())
    
    total_cell = TableCell(stylename=title_style)
    total_cell.addElement(P(text=f"Total: {len(rows)}"))
    title_row.addElement(total_cell)
    table.addElement(title_row)

    header_row = TableRow()
    for header in headers:
        cell = TableCell(stylename=header_style)
        cell.addElement(P(text=str(header)))
        header_row.addElement(cell)
    table.addElement(header_row)

    for row in rows:
        tr = TableRow()
        for col_idx, value in enumerate(row):
            cell_style = data_style
            cell_kwargs = {}

            if isinstance(value, datetime):
                value = value.time() if value.date() == date(1899, 12, 30) else value

            if isinstance(value, datetime):
                cell_style = date_cell_style
                cell_kwargs = {
                    "valuetype": "date",
                    "datevalue": _ods_date_value(value.date()),
                }
                display_text = _format_csv_value(value)
            elif isinstance(value, date):
                cell_style = date_cell_style
                cell_kwargs = {
                    "valuetype": "date",
                    "datevalue": _ods_date_value(value),
                }
                display_text = _format_csv_value(value)
            elif isinstance(value, time):
                cell_style = time_cell_style
                cell_kwargs = {
                    "valuetype": "time",
                    "timevalue": _ods_time_value(value),
                }
                display_text = _format_csv_value(value)
            else:
                display_text = _format_csv_value(value)

            cell = TableCell(stylename=cell_style, **cell_kwargs)
            cell.addElement(P(text=display_text))
            tr.addElement(cell)
        table.addElement(tr)

    doc.spreadsheet.addElement(table)
    doc.save(path)


def _find_soffice() -> str | None:
    env_path = os.environ.get("LIBREOFFICE_PATH")
    if env_path and os.path.isfile(env_path):
        return env_path

    path_exe = shutil.which("soffice") or shutil.which("soffice.exe")
    if path_exe:
        return path_exe

    candidates = [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ]
    for candidate in candidates:
        if os.path.isfile(candidate):
            return candidate
    return None


def _convert_xlsx_to_ods(soffice_path: str, xlsx_path: str, ods_path: str) -> bool:
    out_dir = os.path.dirname(ods_path)
    try:
        result = subprocess.run(
            [
                soffice_path,
                "--headless",
                "--nologo",
                "--nolockcheck",
                "--nodefault",
                "--nofirststartwizard",
                "--convert-to",
                "ods",
                "--outdir",
                out_dir,
                xlsx_path,
            ],
            check=False,
            capture_output=True,
            text=True,
        )
        if result.returncode != 0:
            logging.warning("Falha ao converter para ODS: %s", result.stderr)
            return False

        base_name = os.path.splitext(os.path.basename(xlsx_path))[0]
        generated = os.path.join(out_dir, f"{base_name}.ods")
        if os.path.abspath(generated) != os.path.abspath(ods_path) and os.path.exists(generated):
            if os.path.exists(ods_path):
                os.remove(ods_path)
            os.replace(generated, ods_path)
        return os.path.exists(ods_path)
    except Exception:
        logging.exception("Erro ao converter XLSX para ODS")
        return False


def _cleanup_old_outputs(root_dir: str, target: date) -> None:
    pattern = os.path.join(root_dir, "solicitacao de pickings *.?")
    for path in glob.glob(pattern + "xlsx") + glob.glob(pattern + "csv") + glob.glob(pattern + "ods"):
        base = os.path.basename(path)
        date_part = base.replace("solicitacao de pickings ", "").split(".")[0]
        try:
            file_date = datetime.strptime(date_part, "%d-%m-%Y").date()
        except ValueError:
            continue
        if file_date < target:
            try:
                os.remove(path)
                logging.info("Arquivo antigo removido: %s", path)
            except OSError:
                logging.warning("Falha ao remover arquivo antigo: %s", path)


def _prepare_output_path(path: str) -> str:
    if not os.path.exists(path):
        return path

    try:
        os.remove(path)
        return path
    except PermissionError:
        base, ext = os.path.splitext(path)
        stamp = datetime.now().strftime("%H%M%S")
        alt_path = f"{base} ({stamp}){ext}"
        logging.warning("Arquivo em uso. Usando nome alternativo: %s", alt_path)
        return alt_path
    except OSError:
        base, ext = os.path.splitext(path)
        stamp = datetime.now().strftime("%H%M%S")
        alt_path = f"{base} ({stamp}){ext}"
        logging.warning("Falha ao remover arquivo. Usando nome alternativo: %s", alt_path)
        return alt_path


def main() -> None:
    logging.info("Inicio da execucao")
    root_dir = os.path.dirname(os.path.abspath(__file__))
    input_path = _find_input_workbook(root_dir)
    logging.info("Arquivo de entrada: %s", input_path)

    wb = load_workbook(input_path, data_only=True)
    ws = wb.worksheets[0]
    if ws.auto_filter:
        ws.auto_filter.ref = None
        for row_idx in range(1, (ws.max_row or 0) + 1):
            ws.row_dimensions[row_idx].hidden = False

    all_headers = [
        ws.cell(row=HEADER_ROW, column=col).value for col in range(1, 35)
    ]

    origin_idx = _find_header_index(all_headers, "ORIGEM")
    date_idx = _find_header_index(all_headers, "DATA_SAIDA")

    output_headers = OUTPUT_HEADERS

    target = _target_date(datetime.now())
    logging.info("Data alvo: %s", target.strftime("%d/%m/%Y"))
    _cleanup_old_outputs(root_dir, target)

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
    ods_path = os.path.join(root_dir, f"{base_name}.ods")

    excel_path = _prepare_output_path(excel_path)
    csv_path = _prepare_output_path(csv_path)
    ods_path = _prepare_output_path(ods_path)
    logging.info("Saida XLSX: %s", excel_path)
    logging.info("Saida CSV: %s", csv_path)
    logging.info("Saida ODS: %s", ods_path)

    out_wb, out_ws = _load_output_template(root_dir)
    num_cols = len(OUTPUT_COLUMNS)

    out_ws.merge_cells(f"A1:{get_column_letter(num_cols - 1)}1")
    out_ws.cell(row=1, column=1, value="CDV E CABREUVA")
    out_ws.cell(row=1, column=num_cols, value=f"Total: {len(filtered_rows)}")
    
    thin_border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin"),
    )
    
    for col in range(1, num_cols + 1):
        cell = out_ws.cell(row=1, column=col)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.font = Font(bold=True, color="FFFFFF", size=11)
        cell.fill = PatternFill(start_color="2E5C8A", end_color="2E5C8A", fill_type="solid")
        cell.border = thin_border

    for col_idx, header in enumerate(output_headers, 1):
        cell = out_ws.cell(row=2, column=col_idx, value=header)
        cell.font = Font(bold=True, color="FFFFFF", size=10)
        cell.fill = PatternFill(start_color="4A90E2", end_color="4A90E2", fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border

    for offset, (_, row_filtered) in enumerate(filtered_rows):
        out_row = 3 + offset
        for col_idx, value in enumerate(row_filtered, 1):
            out_ws.cell(row=out_row, column=col_idx, value=value)

    if filtered_rows:
        data_last_row = 2 + len(filtered_rows)
        
        for row in range(3, data_last_row + 1):
            out_ws.cell(row=row, column=1).number_format = "General"
            out_ws.cell(row=row, column=4).number_format = "dd/mm/yyyy"
            out_ws.cell(row=row, column=5).number_format = "hh:mm"
            out_ws.cell(row=row, column=6).number_format = "hh:mm"
            out_ws.row_dimensions[row].height = 15

        for row in out_ws.iter_rows(min_row=1, max_row=data_last_row, min_col=1, max_col=num_cols):
            for cell in row:
                cell.alignment = Alignment(horizontal="center", vertical="center")

        for row in out_ws.iter_rows(min_row=3, max_row=data_last_row, min_col=1, max_col=num_cols):
            for cell in row:
                cell.border = thin_border
        
        for row in out_ws.iter_rows(min_row=2, max_row=2, min_col=1, max_col=num_cols):
            for cell in row:
                cell.border = thin_border
        
        max_lengths = [len(str(h)) for h in output_headers]
        for row in out_ws.iter_rows(min_row=3, max_row=data_last_row, min_col=1, max_col=num_cols):
            for col_idx, cell in enumerate(row, 1):
                length = len(_format_csv_value(cell.value))
                if length > max_lengths[col_idx - 1]:
                    max_lengths[col_idx - 1] = length
        
        for col_idx, length in enumerate(max_lengths, 1):
            adjusted_width = min(length + 2, 30)
            out_ws.column_dimensions[get_column_letter(col_idx)].width = adjusted_width

    out_wb.save(excel_path)

    with open(csv_path, "w", newline="", encoding="utf-8-sig") as csv_file:
        writer = csv.writer(csv_file, delimiter=";")
        writer.writerow(output_headers)
        for _, row_filtered in filtered_rows:
            writer.writerow([_format_csv_value(value) for value in row_filtered])

    soffice_path = _find_soffice()
    if soffice_path:
        if not _convert_xlsx_to_ods(soffice_path, excel_path, ods_path):
            _write_ods(ods_path, output_headers, [row for _, row in filtered_rows])
    else:
        logging.warning("LibreOffice nao encontrado; tentando gerar ODS via odfpy")
        _write_ods(ods_path, output_headers, [row for _, row in filtered_rows])

    print(f"Entrada: {os.path.basename(input_path)}")
    print(f"Data alvo: {target.strftime('%d/%m/%Y')}")
    print(f"Linhas filtradas: {len(filtered_rows)}")
    print(f"Excel: {excel_path}")
    print(f"CSV: {csv_path}")
    print(f"ODS: {ods_path}")
    logging.info("Linhas filtradas: %s", len(filtered_rows))
    logging.info("Execucao concluida")


if __name__ == "__main__":
    log_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "log.txt")
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=[logging.FileHandler(log_path, mode="w", encoding="utf-8"), logging.StreamHandler(sys.stdout)],
    )
    
    # Check license and authorization
    print("\n" + "="*80)
    print("AutoPickingPy - Licensed Software")
    print("="*80 + "\n")
    
    if not check_license_and_authorize():
        print("\n[FATAL] Authorization failed. Application will not run.")
        print("Contact Clebson Luan Alves da Silva for licensing information.")
        sys.exit(1)
    
    print("\n[SUCCESS] Authorization confirmed. Starting application...\n")
    
    try:
        main()
    except Exception:
        logging.exception("Erro inesperado")
        raise

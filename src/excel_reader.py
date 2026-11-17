import os
import openpyxl
import warnings

# Suppress openpyxl user warnings about extensions
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

class ExcelReader:
    def __init__(self, excel_path: str):
        self.excel_path = excel_path
        self.records = [] # list of dicts
        self.of_map = {} # normalized OF -> record

    def load_data(self):
        """Loads data from the Excel file safely in read-only mode."""
        if not os.path.exists(self.excel_path):
            raise FileNotFoundError(f"No se encuentra el archivo Excel en: {self.excel_path}")

        records = []
        of_map = {}

        wb = openpyxl.load_workbook(self.excel_path, data_only=True, read_only=True)
        sheet = None
        for name in ["Listado", "LISTADO", "listado"]:
            if name in wb.sheetnames:
                sheet = wb[name]
                break
        if sheet is None:
            sheet = wb.active

        # Row 1 is header
        # E = index 4 (0-based) -> Pedido ODOO
        # G = index 6 (0-based) -> SB...
        # H = index 7 (0-based) -> Nº OF
        # I = index 8 (0-based) -> Descripción
        # C = index 2 (0-based) -> Cliente
        for row_idx, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=2):
            if not row or not any(row):
                continue
            
            of_raw = row[7] if len(row) > 7 else None
            if not of_raw:
                continue

            of_str = str(of_raw).strip()
            if not of_str or of_str.upper() in ["PTE", "PENDIENTE", "NONE"]:
                continue

            col_e = str(row[4]).strip() if len(row) > 4 and row[4] is not None else ""
            col_g = str(row[6]).strip() if len(row) > 6 and row[6] is not None else ""
            col_i = str(row[8]).strip() if len(row) > 8 and row[8] is not None else ""
            col_c = str(row[2]).strip() if len(row) > 2 and row[2] is not None else ""

            record = {
                "row": row_idx,
                "of": of_str,
                "of_clean": of_str.replace("/", "-").replace("\\", "-").strip(),
                "col_e": col_e,
                "col_g": col_g,
                "col_i": col_i,
                "cliente": col_c,
                "display": f"{of_str} | {col_e} | {col_i[:35]}"
            }
            records.append(record)
            
            # Map various keys for quick lookup
            of_map[of_str.upper()] = record
            of_map[record["of_clean"].upper()] = record
            # Also map just the numeric/suffix part if it looks like AYC/MO/00567 -> 00567 and 567
            parts = of_str.replace("\\", "/").split("/")
            if len(parts) > 1:
                last_part = parts[-1].strip()
                if last_part:
                    of_map[last_part.upper()] = record
                    if last_part.isdigit():
                        of_map[str(int(last_part))] = record

        wb.close()
        self.records = records
        self.of_map = of_map
        return len(records)

    def find_of(self, query: str):
        """Finds record by exact or partial OF / query."""
        if not query:
            return None
        q = str(query).strip().upper()
        
        # Exact match
        if q in self.of_map:
            return self.of_map[q]
        
        # Cleaned match (replace / with -)
        q_clean = q.replace("/", "-").replace("\\", "-")
        if q_clean in self.of_map:
            return self.of_map[q_clean]

        # Partial substring in OF or Odoo or Desc
        for rec in self.records:
            if q in rec["of"].upper() or q in rec["col_e"].upper():
                return rec
        
        return None

    def search_suggestions(self, query: str, limit: int = 15):
        """Returns autocomplete suggestions based on query."""
        if not query:
            return [r["of"] for r in self.records[:limit]]
        q = query.strip().upper()
        matches = []
        for r in self.records:
            if q in r["of"].upper() or q in r["col_e"].upper() or q in r["cliente"].upper() or q in r["col_i"].upper():
                matches.append(r["of"])
                if len(matches) >= limit:
                    break
        return matches

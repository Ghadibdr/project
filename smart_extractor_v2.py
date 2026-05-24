import anthropic
import openpyxl
import pandas as pd
import hashlib
import json
import sqlite3
import os
import io
import base64
import subprocess
import sys
import tempfile

client = anthropic.Anthropic()

# ============================================================
# KNOWLEDGE BASE - يحفظ كل طريقة استخراج تعلمها النظام
# ============================================================

class KnowledgeBase:
    def __init__(self, db_path="extraction_kb.db"):
        self.db_path = db_path
        self._init()

    def _init(self):
        with sqlite3.connect(self.db_path) as conn:
            conn.execute("""
                CREATE TABLE IF NOT EXISTS methods (
                    fingerprint   TEXT PRIMARY KEY,
                    code          TEXT NOT NULL,
                    file_type     TEXT,
                    description   TEXT,
                    used          INTEGER DEFAULT 1,
                    created_at    TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            """)

    def get(self, fingerprint):
        with sqlite3.connect(self.db_path) as conn:
            row = conn.execute(
                "SELECT code FROM methods WHERE fingerprint=?", (fingerprint,)
            ).fetchone()
            if row:
                conn.execute(
                    "UPDATE methods SET used=used+1 WHERE fingerprint=?", (fingerprint,)
                )
            return row[0] if row else None

    def save(self, fingerprint, code, file_type="", description=""):
        with sqlite3.connect(self.db_path) as conn:
            conn.execute(
                "INSERT OR REPLACE INTO methods (fingerprint,code,file_type,description) VALUES(?,?,?,?)",
                (fingerprint, code, file_type, description)
            )


# ============================================================
# FINGERPRINTER - يعمل بصمة لكل شكل بيانات
# ============================================================

class Fingerprinter:

    def excel_section(self, ws, start_row, end_row):
        """بصمة لقسم معين داخل sheet."""
        headers, merge_patterns, type_patterns = [], [], []

        for row in ws.iter_rows(min_row=start_row, max_row=min(start_row + 6, end_row)):
            for cell in row:
                if cell.font and cell.font.bold and cell.value:
                    headers.append(str(cell.value)[:40])

        for m in ws.merged_cells.ranges:
            if m.min_row >= start_row and m.max_row <= end_row:
                merge_patterns.append(f"r{m.max_row - m.min_row}c{m.max_col - m.min_col}")

        for row in ws.iter_rows(min_row=start_row + 1, max_row=min(start_row + 4, end_row)):
            types = [type(c.value).__name__ for c in row if c.value is not None]
            if types:
                type_patterns.append(types)

        features = {
            "col_count": ws.max_column,
            "headers": sorted(headers),
            "merges": sorted(merge_patterns),
            "types": type_patterns[:2],
        }
        fp = hashlib.md5(json.dumps(features, sort_keys=True).encode()).hexdigest()
        return fp, features

    def detect_sections(self, ws):
        """يكتشف الأقسام المختلفة داخل sheet واحدة."""
        sections, start = [], 1

        for r in range(1, ws.max_row + 1):
            row_cells = ws[r]
            empty = all(c.value is None for c in row_cells)
            is_header = any(c.font and c.font.bold and c.value for c in row_cells)

            if r > start + 2 and (is_header or empty):
                if r - start > 2:
                    sections.append((start, r - 1))
                start = r + 1 if empty else r

        if ws.max_row - start > 1:
            sections.append((start, ws.max_row))

        return sections or [(1, ws.max_row)]

    def csv_file(self, filepath):
        df = pd.read_csv(filepath, nrows=5)
        features = {
            "columns": list(df.columns),
            "dtypes": dict(df.dtypes.astype(str)),
        }
        fp = hashlib.md5(json.dumps(features, sort_keys=True).encode()).hexdigest()
        return fp, features


# ============================================================
# VISION GENERATOR - يفهم الشكل الجديد ويكتب كود
# ============================================================

class VisionGenerator:

    def _section_to_rich_text(self, ws, start_row, end_row):
        lines = []
        for row in ws.iter_rows(min_row=start_row, max_row=min(end_row, start_row + 60)):
            for cell in row:
                if cell.value is None:
                    continue
                bold = bool(cell.font and cell.font.bold)
                filled = False
                try:
                    if cell.fill and cell.fill.fgColor:
                        filled = cell.fill.fgColor.rgb not in ("00000000", "FFFFFFFF")
                except Exception:
                    pass
                lines.append(
                    f"[{cell.row},{cell.column}] {repr(str(cell.value)[:50])} "
                    f"bold={bold} filled={filled}"
                )
        return "\n".join(lines)

    def _pdf_to_base64(self, filepath, page=0):
        try:
            from pdf2image import convert_from_path
            pages = convert_from_path(filepath, dpi=150)
            if page < len(pages):
                buf = io.BytesIO()
                pages[page].save(buf, format="PNG")
                return base64.standard_b64encode(buf.getvalue()).decode()
        except Exception:
            pass
        return None

    def generate_code(self, filepath, section_text, fp_features, section_info):
        abs_path = os.path.abspath(filepath)
        ext = os.path.splitext(filepath)[1].lower()

        # لو PDF → نجرب Vision
        img_b64 = None
        if ext == ".pdf":
            img_b64 = self._pdf_to_base64(filepath)

        if img_b64:
            content = [
                {
                    "type": "image",
                    "source": {"type": "base64", "media_type": "image/png", "data": img_b64}
                },
                {
                    "type": "text",
                    "text": f"""You are a data extraction expert.
Study this document image carefully and understand its exact structure.

FILE PATH: {abs_path}

Write Python code that:
1. Opens the file from the path above
2. Extracts ALL meaningful data based on the structure you see
3. Organizes it into a clean Python dict or list of dicts
4. Prints result: print(json.dumps(result, ensure_ascii=False, indent=2, default=str))

Return ONLY executable Python code. No markdown, no explanation."""
                }
            ]
        else:
            # للباقي → rich text representation
            content = f"""You are a data extraction expert.
Below is a detailed structural map of a file section.
Bold cells = headers/labels. Merged cells = grouped data.
Analyze the pattern and understand what data is where.

FILE PATH: {abs_path}
FILE TYPE: {ext}
SECTION INFO: rows {section_info.get('start_row', '?')} to {section_info.get('end_row', '?')}
STRUCTURAL FEATURES: {json.dumps(fp_features, ensure_ascii=False)}

CELL MAP:
{section_text}

Write Python code that:
1. Opens {abs_path}
2. Extracts ALL data from this section correctly
3. Handles merged cells and irregular layout
4. Returns clean dict/list
5. Prints: print(json.dumps(result, ensure_ascii=False, indent=2, default=str))

Return ONLY executable Python code. No markdown, no explanation."""

        response = client.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=3000,
            messages=[{"role": "user", "content": content}]
        )
        return response.content[0].text.strip()

    def fix_code(self, code, error, section_text, fp_features):
        response = client.messages.create(
            model="claude-haiku-4-5-20251001",
            max_tokens=2000,
            messages=[{"role": "user", "content": f"""Fix this Python code.

ERROR: {error}

FAILED CODE:
{code}

STRUCTURE CONTEXT:
{json.dumps(fp_features)}

Return ONLY fixed Python code. No markdown."""}]
        )
        return response.content[0].text.strip()


# ============================================================
# SANDBOX - يشغل الكود في بيئة معزولة
# ============================================================

def run_sandbox(code, timeout=60):
    code = code.strip()
    if code.startswith("```"):
        lines = [l for l in code.split("\n") if not l.startswith("```")]
        code = "\n".join(lines)

    with tempfile.NamedTemporaryFile(mode="w", suffix=".py", delete=False, encoding="utf-8") as f:
        f.write(code)
        tmp = f.name

    try:
        r = subprocess.run(
            [sys.executable, tmp],
            capture_output=True, text=True, timeout=timeout
        )
        return r.stdout.strip(), r.stderr.strip(), r.returncode
    except subprocess.TimeoutExpired:
        return "", "Timeout (60s)", -1
    finally:
        try:
            os.unlink(tmp)
        except Exception:
            pass


# ============================================================
# SMART EXTRACTOR - المحرك الرئيسي
# ============================================================

class SmartExtractor:
    def __init__(self, kb_path="extraction_kb.db"):
        self.kb = KnowledgeBase(kb_path)
        self.fp = Fingerprinter()
        self.vision = VisionGenerator()

    def _run_with_retry(self, code, section_text, fp_features, max_retries=2):
        for attempt in range(max_retries + 1):
            out, err, status = run_sandbox(code)
            if status == 0 and out:
                return out, code, True
            if attempt < max_retries and err:
                print(f"      error → asking LLM to fix (attempt {attempt+1})")
                code = self.vision.fix_code(code, err, section_text, fp_features)
        return None, code, False

    def _process_section(self, filepath, section_text, fp_hash, fp_features, section_info, file_type):
        cached = self.kb.get(fp_hash)

        if cached:
            print(f"      ✓ Known structure (fp:{fp_hash[:8]}…) → using cached code")
            code = cached
        else:
            print(f"      ✗ New structure (fp:{fp_hash[:8]}…) → calling Vision/LLM")
            code = self.vision.generate_code(filepath, section_text, fp_features, section_info)

        out, final_code, success = self._run_with_retry(code, section_text, fp_features)

        if success:
            if not cached:
                desc = str(fp_features.get("headers", ""))[:100]
                self.kb.save(fp_hash, final_code, file_type, desc)
                print(f"      ✓ Saved to Knowledge Base")
            try:
                return json.loads(out)
            except Exception:
                return out
        else:
            print(f"      ✗ Extraction failed for this section")
            return None

    def extract_excel(self, filepath):
        wb = openpyxl.load_workbook(filepath, data_only=True)
        all_results = {}

        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            sections = self.fp.detect_sections(ws)
            print(f"\n  Sheet '{sheet_name}': {len(sections)} section(s) detected")
            sheet_data = []

            for i, (s_row, e_row) in enumerate(sections):
                print(f"\n    Section {i+1} (rows {s_row}–{e_row})")
                fp_hash, fp_features = self.fp.excel_section(ws, s_row, e_row)

                section_text = ""
                cached = self.kb.get(fp_hash)
                if not cached:
                    section_text = self._excel_section_text(ws, s_row, e_row)

                result = self._process_section(
                    filepath, section_text, fp_hash, fp_features,
                    {"start_row": s_row, "end_row": e_row}, "xlsx"
                )
                if result:
                    sheet_data.append(result)

            all_results[sheet_name] = sheet_data

        return all_results

    def extract_csv(self, filepath):
        print(f"\n  CSV file")
        fp_hash, fp_features = self.fp.csv_file(filepath)

        section_text = ""
        if not self.kb.get(fp_hash):
            section_text = pd.read_csv(filepath, nrows=30).to_string()

        return self._process_section(
            filepath, section_text, fp_hash, fp_features, {}, "csv"
        )

    def extract_json(self, filepath):
        print(f"\n  JSON file")
        with open(filepath, encoding="utf-8") as f:
            raw = f.read(3000)
        fp_hash = hashlib.md5(raw[:500].encode()).hexdigest()
        return self._process_section(
            filepath, raw, fp_hash, {}, {}, "json"
        )

    def extract_pdf(self, filepath):
        print(f"\n  PDF file → using Vision")
        fp_hash = hashlib.md5(open(filepath, "rb").read(1000)).hexdigest()
        return self._process_section(
            filepath, "", fp_hash, {}, {"page": 0}, "pdf"
        )

    def extract(self, filepath):
        if not os.path.exists(filepath):
            raise FileNotFoundError(filepath)

        ext = os.path.splitext(filepath)[1].lower()
        print(f"\n{'='*55}")
        print(f"SmartExtractor: {os.path.basename(filepath)}")
        print(f"{'='*55}")

        dispatch = {
            ".xlsx": self.extract_excel,
            ".xls":  self.extract_excel,
            ".xlsm": self.extract_excel,
            ".csv":  self.extract_csv,
            ".json": self.extract_json,
            ".pdf":  self.extract_pdf,
        }

        handler = dispatch.get(ext, self.extract_json)
        return handler(filepath)

    @staticmethod
    def _excel_section_text(ws, start_row, end_row):
        lines = []
        for row in ws.iter_rows(min_row=start_row, max_row=min(end_row, start_row + 60)):
            for cell in row:
                if cell.value is None:
                    continue
                bold = bool(cell.font and cell.font.bold)
                lines.append(f"[{cell.row},{cell.column}] {repr(str(cell.value)[:50])} bold={bold}")
        return "\n".join(lines)


# ============================================================
# ENTRY POINT
# ============================================================

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python smart_extractor_v2.py <file>")
        sys.exit(1)

    extractor = SmartExtractor(kb_path="extraction_kb.db")
    result = extractor.extract(sys.argv[1])

    print("\n" + "="*55)
    print("RESULT:")
    print("="*55)
    print(json.dumps(result, ensure_ascii=False, indent=2, default=str))

    out_path = os.path.splitext(sys.argv[1])[0] + "_extracted.json"
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2, default=str)
    print(f"\nSaved → {out_path}")

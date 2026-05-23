import anthropic
import openpyxl
import pandas as pd
import subprocess
import tempfile
import os
import json
import sys

client = anthropic.Anthropic()


def represent_excel(filepath, max_rows=80):
    """Convert Excel to rich text that preserves structure for LLM understanding."""
    wb = openpyxl.load_workbook(filepath, data_only=True)
    parts = []

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        parts.append(f"=== Sheet: {sheet_name} ===")
        parts.append(f"Size: {ws.max_row} rows x {ws.max_column} columns")

        if ws.merged_cells.ranges:
            merged = [str(r) for r in ws.merged_cells.ranges]
            parts.append(f"Merged cells: {', '.join(merged)}")

        parts.append("\n[row, col] value | bold | colored")
        parts.append("-" * 40)

        for row in ws.iter_rows(min_row=1, max_row=min(max_rows, ws.max_row)):
            for cell in row:
                if cell.value is None:
                    continue
                bold = cell.font.bold if cell.font else False
                colored = False
                if cell.fill and cell.fill.fgColor:
                    try:
                        colored = cell.fill.fgColor.rgb not in ("00000000", "FFFFFFFF")
                    except Exception:
                        pass
                parts.append(f"  [{cell.row},{cell.column}] {repr(cell.value)} | bold={bold} | colored={colored}")

    return "\n".join(parts)


def represent_csv(filepath, max_rows=80):
    """Convert CSV to text with structure hints."""
    df = pd.read_csv(filepath, nrows=max_rows)
    parts = [
        f"CSV File: {os.path.basename(filepath)}",
        f"Shape: {df.shape[0]} rows x {df.shape[1]} columns",
        f"Columns: {list(df.columns)}",
        f"Data types: {df.dtypes.to_dict()}",
        "\nFirst rows:",
        df.to_string(max_rows=20)
    ]
    return "\n".join(str(p) for p in parts)


def represent_json(filepath):
    """Represent JSON structure."""
    with open(filepath) as f:
        data = json.load(f)
    preview = json.dumps(data, ensure_ascii=False, indent=2)[:3000]
    return f"JSON File:\n{preview}"


def get_data_representation(filepath):
    """Auto-detect file type and return rich text representation."""
    ext = os.path.splitext(filepath)[1].lower()
    if ext in (".xlsx", ".xls"):
        return represent_excel(filepath)
    elif ext == ".csv":
        return represent_csv(filepath)
    elif ext == ".json":
        return represent_json(filepath)
    else:
        with open(filepath) as f:
            return f.read(5000)


def generate_extraction_code(data_repr, filepath):
    """Ask LLM to generate extraction code based on data structure."""
    abs_path = os.path.abspath(filepath)
    ext = os.path.splitext(filepath)[1].lower()

    response = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=3000,
        messages=[{
            "role": "user",
            "content": f"""You are an expert data engineer.
Below is a detailed structural representation of a data file.
Study the structure carefully — cell positions, merged cells, bold headers, and patterns.

FILE PATH: {abs_path}
FILE TYPE: {ext}

=== DATA STRUCTURE ===
{data_repr}
=== END ===

Write Python code that:
1. Reads the file from the exact path above
2. Understands the layout from the structure shown
3. Extracts all meaningful data correctly
4. Organizes it into a clean Python dict or list of dicts
5. Prints the final result as formatted JSON

Rules:
- Return ONLY executable Python code, no markdown, no explanation
- Use openpyxl for xlsx, pandas for csv, json module for json
- Handle merged cells and irregular structures properly
- The output must be valid JSON printed to stdout"""
        }]
    )
    return response.content[0].text.strip()


def clean_code(code):
    """Remove markdown code blocks if LLM added them."""
    if code.startswith("```"):
        lines = code.split("\n")
        lines = [l for l in lines if not l.startswith("```")]
        code = "\n".join(lines)
    return code


def run_in_sandbox(code, timeout=30):
    """Run generated code in isolated subprocess."""
    with tempfile.NamedTemporaryFile(mode="w", suffix=".py", delete=False) as f:
        f.write(code)
        tmp = f.name

    try:
        result = subprocess.run(
            [sys.executable, tmp],
            capture_output=True,
            text=True,
            timeout=timeout
        )
        return result.stdout.strip(), result.stderr.strip()
    except subprocess.TimeoutExpired:
        return "", "Timeout: code took too long"
    finally:
        os.unlink(tmp)


def retry_with_error(data_repr, filepath, code, error, attempt):
    """Ask LLM to fix the code given the error."""
    response = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=3000,
        messages=[{
            "role": "user",
            "content": f"""This Python code failed with an error. Fix it.

=== ORIGINAL DATA STRUCTURE ===
{data_repr}
=== END ===

=== FAILED CODE ===
{code}
=== END ===

=== ERROR ===
{error}
=== END ===

Return ONLY the fixed Python code, no markdown, no explanation."""
        }]
    )
    return response.content[0].text.strip()


def extract(filepath, max_retries=2):
    """Main extraction pipeline."""
    print(f"\n[1] Analyzing structure of: {filepath}")
    data_repr = get_data_representation(filepath)

    print("[2] Asking LLM to generate extraction code...")
    code = clean_code(generate_extraction_code(data_repr, filepath))

    for attempt in range(max_retries + 1):
        print(f"[3] Running code in sandbox (attempt {attempt + 1})...")
        output, error = run_in_sandbox(code)

        if output and not error:
            print("[4] Success!\n")
            print("=== EXTRACTED DATA ===")
            print(output)
            return output

        if error:
            print(f"    Error: {error[:200]}")
            if attempt < max_retries:
                print("[3] Asking LLM to fix the error...")
                code = clean_code(retry_with_error(data_repr, filepath, code, error, attempt))
            else:
                print("[!] Failed after all retries.")
                print("Last error:", error)

    return None


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python smart_extractor.py <file_path>")
        print("Supported: .xlsx, .xls, .csv, .json")
        sys.exit(1)

    filepath = sys.argv[1]
    if not os.path.exists(filepath):
        print(f"File not found: {filepath}")
        sys.exit(1)

    extract(filepath)

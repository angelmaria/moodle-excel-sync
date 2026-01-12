from __future__ import annotations

import argparse
import unicodedata
from datetime import datetime
from pathlib import Path

import pandas as pd


def _norm_key(value: str) -> str:
    s = (value or "").strip().lower()
    s = "".join(c for c in unicodedata.normalize("NFKD", s) if not unicodedata.combining(c))
    s = " ".join(s.replace("\n", " ").split())
    return s


def _pick_email_column(columns) -> str | None:
    normalized = {c: _norm_key(str(c)) for c in columns}

    for target in ("correo", "email", "e-mail", "direccion de correo", "dirección de correo"):
        t = _norm_key(target)
        for c, n in normalized.items():
            if n == t:
                return str(c)

    for c, n in normalized.items():
        if "correo" in n or "email" in n:
            return str(c)

    return None


def _emails_from_excel(path: Path) -> tuple[set[str], str | None]:
    df = pd.read_excel(path)
    col = _pick_email_column(df.columns)
    if not col:
        return set(), None

    s = df[col].dropna().astype(str).map(lambda x: x.strip().lower())
    s = s[s != ""]
    return set(s.tolist()), col


def _build_arg_parser() -> argparse.ArgumentParser:
    base_dir = Path(__file__).resolve().parent

    p = argparse.ArgumentParser(description="Chequea emails en Excel vs export CSV de Moodle")
    p.add_argument(
        "--excel-dir",
        type=Path,
        default=base_dir / "excel",
        help="Carpeta con .xlsx (por defecto: ./excel)",
    )
    p.add_argument(
        "--csv",
        type=Path,
        default=base_dir / "Usuarios_12_enero_2026.csv",
        help="CSV exportado de Moodle con columna 'email' (por defecto: ./Usuarios_12_enero_2026.csv)",
    )
    p.add_argument(
        "--log-dir",
        type=Path,
        default=base_dir / "logs",
        help="Carpeta donde guardar el log (por defecto: ./logs)",
    )
    p.add_argument(
        "--max-sample",
        type=int,
        default=20,
        help="Cuántos emails faltantes mostrar por archivo en consola (por defecto: 20)",
    )
    return p


def main() -> int:
    args = _build_arg_parser().parse_args()

    csv_path: Path = args.csv
    excel_dir: Path = args.excel_dir
    log_dir: Path = args.log_dir
    max_sample: int = args.max_sample

    log_dir.mkdir(exist_ok=True)
    log_path = log_dir / f"final_check_excel_vs_csv__{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"

    df_csv = pd.read_csv(csv_path)
    if "email" not in df_csv.columns:
        raise SystemExit(f"No encuentro columna 'email' en {csv_path.name}: {list(df_csv.columns)}")
    emails_csv = set(df_csv["email"].dropna().astype(str).map(lambda x: x.strip().lower()).tolist())

    excel_files = sorted(excel_dir.glob("*.xlsx")) if excel_dir.exists() else []

    lines: list[str] = []
    lines.append(f"CSV (Moodle export): {csv_path} -> {len(emails_csv)} emails")
    lines.append(f"Excel dir: {excel_dir} -> {len(excel_files)} archivos .xlsx")
    lines.append("")

    print(f"CSV (Moodle export): {len(emails_csv)} emails")
    print(f"Excel dir: {excel_dir} ({len(excel_files)} .xlsx)")

    any_missing = False
    for path in excel_files:
        try:
            emails_xlsx, col = _emails_from_excel(path)
        except Exception as e:
            print(f"- {path.name}: ERROR leyendo ({type(e).__name__}: {e})")
            lines.append(f"- {path.name}: ERROR leyendo ({type(e).__name__}: {e})")
            continue

        if col is None:
            print(f"- {path.name}: NO pude detectar columna de email")
            lines.append(f"- {path.name}: NO pude detectar columna de email")
            continue

        missing = sorted(e for e in emails_xlsx if e not in emails_csv)
        print(f"- {path.name}: {len(emails_xlsx)} emails (col='{col}') -> faltan en CSV: {len(missing)}")
        lines.append(f"- {path.name}: {len(emails_xlsx)} emails (col='{col}') -> faltan en CSV: {len(missing)}")

        if missing:
            any_missing = True
            for e in missing[:max_sample]:
                print(f"  - {e}")
                lines.append(f"  - {e}")

            if len(missing) > max_sample:
                print(f"  ... +{len(missing) - max_sample} más")
                lines.append(f"  ... +{len(missing) - max_sample} más")

    lines.append("")
    lines.append(
        "OK: todos los emails de Excel están en el CSV"
        if not any_missing
        else "ATENCIÓN: hay emails de Excel que no están en el CSV"
    )

    log_path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    print(f"Log escrito en: {log_path}")

    return 0 if not any_missing else 2


if __name__ == "__main__":
    raise SystemExit(main())

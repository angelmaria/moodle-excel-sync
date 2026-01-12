from __future__ import annotations

import argparse
from pathlib import Path
from datetime import datetime
import pandas as pd

COL_APELLIDOS = "Apellidos"
COL_NOMBRE = "Nombre"
COL_CORREO = "Correo"
COL_TELEFONO = "Número de teléfono"
COL_PAIS = "País/región"
COL_USUARIO = "Usuario"
COL_CONTRASENA = "Contraseña"

OUTPUT_COLUMNS = [
    COL_APELLIDOS,
    COL_NOMBRE,
    COL_CORREO,
    COL_TELEFONO,
    COL_PAIS,
    COL_USUARIO,
    COL_CONTRASENA,
]


def normalize_email(value) -> str | None:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    s = str(value).strip().lower()
    if not s:
        return None
    return s


def email_local_part(email: str) -> str:
    email = (email or "").strip()
    if "@" not in email:
        raise ValueError(f"Email inválido (sin @): {email!r}")
    return email.split("@", 1)[0].strip()


def first_name(nombre: str) -> str:
    nombre = (nombre or "").strip()
    if not nombre:
        raise ValueError("Nombre vacío, no puedo generar contraseña.")
    return nombre.split()[0]


def read_emails_from_excel(path: Path) -> set[str]:
    df = pd.read_excel(path)
    if COL_CORREO not in df.columns:
        raise KeyError(f"No encuentro la columna {COL_CORREO!r} en {path.name}. Columnas: {list(df.columns)}")
    emails = set()
    for v in df[COL_CORREO].tolist():
        e = normalize_email(v)
        if e:
            emails.add(e)
    return emails


def main() -> None:
    parser = argparse.ArgumentParser(
        description=(
            "Genera un Excel con los registros faltantes (por email) y añade Usuario/Contraseña. "
            "Usuario = parte antes del @. Contraseña = primer nombre + '+A1+-'. "
            "El Excel de salida se ordena para ser compatible con moodle_excel_sync.py."
        )
    )
    parser.add_argument(
        "--input",
        default="registro_curso_amor_sexualidad4.xlsx",
        help="Excel de entrada (por defecto: registro_curso_amor_sexualidad4.xlsx)",
    )
    parser.add_argument(
        "--compare",
        default=["registro_curso_amor_sexualidad2_rellenado.xlsx", "registro_curso_amor_sexualidad3_faltantes_rellenado.xlsx"],
        nargs="*",
        help="Lista de excels ya procesados/subidos para excluir emails (por defecto: 2_rellenado y 3_faltantes)",
    )
    parser.add_argument(
        "--output",
        default=None,
        help="Excel de salida. Si no se indica: <input_stem>_faltantes_rellenado.xlsx",
    )

    args = parser.parse_args()

    base_dir = Path(__file__).resolve().parent
    excel_dir = base_dir / "excel"
    input_xlsx = (excel_dir / args.input).resolve()
    compare_xlsx = [(excel_dir / p).resolve() for p in args.compare]
    output_xlsx = (excel_dir / args.output).resolve() if args.output else (excel_dir / f"{Path(args.input).stem}_faltantes_rellenado.xlsx")

    log_dir = base_dir / "logs"
    log_dir.mkdir(exist_ok=True)
    run_ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_file = log_dir / f"log_prepare_excel__{input_xlsx.stem}__{run_ts}.txt"

    def log(line: str) -> None:
        print(line)
        try:
            with log_file.open("a", encoding="utf-8") as f:
                f.write(line + "\n")
        except Exception:
            pass

    # Reiniciar log
    try:
        log_file.write_text("", encoding="utf-8")
    except Exception:
        pass

    log("Preparando faltantes por email")
    log(f"Input: {input_xlsx.name}")
    log(f"Output: {output_xlsx.name}")

    df_in = pd.read_excel(input_xlsx)
    log(f"Registros input: {len(df_in)}")

    for required in (COL_APELLIDOS, COL_NOMBRE, COL_CORREO):
        if required not in df_in.columns:
            raise KeyError(f"No encuentro la columna {required!r} en {input_xlsx.name}. Columnas: {list(df_in.columns)}")

    # Emails ya existentes (tratados)
    existing_emails: set[str] = set()
    for p in compare_xlsx:
        if not p.exists():
            log(f"⚠ No existe compare file: {p.name} (se ignora)")
            continue
        emails = read_emails_from_excel(p)
        log(f"Compare: {p.name} -> {len(emails)} emails")
        existing_emails |= emails
    log(f"Total emails existentes (unión): {len(existing_emails)}")

    df = df_in.copy()
    df["__email_norm__"] = df[COL_CORREO].apply(normalize_email)

    dup_mask = df["__email_norm__"].notna() & df["__email_norm__"].duplicated(keep="first")
    dup_emails = sorted(set(df.loc[dup_mask, "__email_norm__"].tolist()))
    if dup_emails:
        log(f"⚠ Emails duplicados dentro de {input_xlsx.name}: {len(dup_emails)} (se conserva la primera aparición)")
        for e in dup_emails[:50]:
            log(f"  - {e}")
        if len(dup_emails) > 50:
            log(f"  ... +{len(dup_emails) - 50} más")

    df = df[df["__email_norm__"].notna()].drop_duplicates(subset=["__email_norm__"], keep="first")
    df_f = df[~df["__email_norm__"].isin(existing_emails)].copy()
    log(f"Faltantes detectados (por email): {len(df_f)}")

    for c in (COL_TELEFONO, COL_PAIS):
        if c not in df_f.columns:
            df_f[c] = ""

    invalid_email = 0
    empty_name = 0

    usuarios = []
    contras = []
    for _, row in df_f.iterrows():
        correo = str(row[COL_CORREO]).strip() if pd.notna(row[COL_CORREO]) else ""
        nombre = str(row[COL_NOMBRE]).strip() if pd.notna(row[COL_NOMBRE]) else ""

        try:
            u = email_local_part(correo)
        except Exception:
            u = ""
            invalid_email += 1

        try:
            c = f"{first_name(nombre)}+A1+-" if nombre else ""
        except Exception:
            c = ""
            empty_name += 1

        usuarios.append(u)
        contras.append(c)

    df_f[COL_USUARIO] = usuarios
    df_f[COL_CONTRASENA] = contras

    if invalid_email:
        log(f"⚠ Emails inválidos (sin @) en faltantes: {invalid_email} (se deja Usuario vacío)")
    if empty_name:
        log(f"⚠ Nombres vacíos/problema para contraseña en faltantes: {empty_name} (se deja Contraseña vacía)")

    df_out = pd.DataFrame({c: df_f[c] if c in df_f.columns else "" for c in OUTPUT_COLUMNS})
    df_out.to_excel(output_xlsx, index=False, sheet_name="Usuarios")

    log(f"OK generado: {output_xlsx.name} ({len(df_out)} filas)")
    log(f"Log preparación: {log_file.name}")


if __name__ == "__main__":
    main()

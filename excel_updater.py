#!/usr/bin/env python3

import argparse
import csv
import os
import re
import sys
from typing import Dict, Optional, Tuple, List

from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.worksheet import Worksheet


def load_worksheet_from_input(input_path: str) -> Tuple[Workbook, Worksheet]:
    """Load an Excel (.xlsx/.xls) or CSV file into an openpyxl Workbook+Worksheet.

    For CSV, creates a new Workbook and fills it with the CSV contents.
    Assumes the first row is the header.
    """
    lower_path = input_path.lower()
    if lower_path.endswith((".xlsx", ".xls")):
        wb = load_workbook(input_path, data_only=True)
        ws = wb.active
        return wb, ws

    if lower_path.endswith(".csv"):
        wb = Workbook()
        ws = wb.active
        with open(input_path, "r", newline="", encoding="utf-8") as f:
            reader = csv.reader(f)
            for row in reader:
                ws.append(row)
        return wb, ws

    raise ValueError(
        f"Formato de entrada no soportado para '{input_path}'. Use .xlsx, .xls o .csv"
    )


def ensure_new_column_position(ws: Worksheet, item_column_name: str, new_column_name: str) -> int:
    """Ensure the new column exists and is placed right after the item column.

    Returns the 1-based index of the new column.
    """
    headers = [str(cell.value) if cell.value is not None else "" for cell in ws[1]]
    if item_column_name not in headers:
        raise KeyError(
            f"La columna de item '{item_column_name}' no existe. Encabezados: {headers}"
        )

    item_idx_one_based = headers.index(item_column_name) + 1

    # If new column already exists, move it next to item_column if needed
    if new_column_name in headers:
        current_idx = headers.index(new_column_name) + 1
        desired_idx = item_idx_one_based + 1
        if current_idx != desired_idx:
            # Insert a temp column and copy values, then delete the original to reposition
            ws.insert_cols(desired_idx)
            for r in range(1, ws.max_row + 1):
                ws.cell(row=r, column=desired_idx).value = (
                    new_column_name if r == 1 else ws.cell(row=r, column=current_idx + 1).value
                )
            # After inserting before, the original index shifts by +1 if it was >= desired_idx
            # Determine the index to delete by finding the other header occurrence
            # Rebuild headers and find the other instance
            headers2 = [str(cell.value) if cell.value is not None else "" for cell in ws[1]]
            # Remove only one occurrence not at desired_idx
            to_delete_idx = None
            for idx, name in enumerate(headers2, start=1):
                if name == new_column_name and idx != desired_idx:
                    to_delete_idx = idx
                    break
            if to_delete_idx is not None:
                ws.delete_cols(to_delete_idx)
        return item_idx_one_based + 1

    # Insert a new column right after item column
    insert_at = item_idx_one_based + 1
    ws.insert_cols(insert_at)
    ws.cell(row=1, column=insert_at).value = new_column_name
    return insert_at


def parse_map_string(map_str: str, pair_sep: str = ";", kv_sep: str = "=") -> Dict[str, str]:
    mapping: Dict[str, str] = {}
    if not map_str:
        return mapping
    for part in [p.strip() for p in map_str.split(pair_sep) if p.strip()]:
        if kv_sep not in part:
            raise ValueError(
                f"Entrada de mapeo inválida '{part}'. Use el formato clave{kv_sep}valor"
            )
        key, value = part.split(kv_sep, 1)
        mapping[key.strip()] = value.strip()
    return mapping


def normalize_text(value: Optional[str], strip_values: bool, case_sensitive: bool) -> str:
    text = "" if value is None else str(value)
    if strip_values:
        text = text.strip()
    if not case_sensitive:
        text = text.lower()
    return text


def matches(value: str, key: str, match_mode: str) -> bool:
    if match_mode == "exact":
        return value == key
    if match_mode == "contains":
        return key in value
    if match_mode == "regex":
        try:
            return re.search(key, value) is not None
        except re.error as exc:
            raise ValueError(f"Expresión regular inválida '{key}': {exc}") from exc
    raise ValueError(f"Modo de coincidencia no soportado: {match_mode}")


def apply_single_mapping(
    ws: Worksheet,
    item_col_idx: int,
    new_col_idx: int,
    key: str,
    value: str,
    match_mode: str,
    case_sensitive: bool,
    strip_values: bool,
    skip_existing: bool,
) -> int:
    key_norm = normalize_text(key, strip_values, case_sensitive)
    updates = 0
    for r in range(2, ws.max_row + 1):
        cell_value = ws.cell(row=r, column=item_col_idx).value
        cell_text = normalize_text(str(cell_value) if cell_value is not None else "", strip_values, case_sensitive)

        if matches(cell_text, key_norm, match_mode):
            target_cell = ws.cell(row=r, column=new_col_idx)
            if skip_existing:
                existing = target_cell.value
                existing_text = "" if existing is None else str(existing)
                if len(existing_text) > 0:
                    continue
            target_cell.value = value
            updates += 1
    return updates


def apply_mapping_dict(
    ws: Worksheet,
    item_col_idx: int,
    new_col_idx: int,
    mapping: Dict[str, str],
    match_mode: str,
    case_sensitive: bool,
    strip_values: bool,
    skip_existing: bool,
) -> Tuple[int, int]:
    total_updates = 0
    keys_with_hits = 0
    for key, mapped_value in mapping.items():
        hits = apply_single_mapping(
            ws,
            item_col_idx,
            new_col_idx,
            key,
            mapped_value,
            match_mode,
            case_sensitive,
            strip_values,
            skip_existing,
        )
        total_updates += hits
        if hits > 0:
            keys_with_hits += 1
    return total_updates, keys_with_hits


def load_mapping_from_file(mapping_file: str, key_col: Optional[str], val_col: Optional[str]) -> Dict[str, str]:
    lower_path = mapping_file.lower()
    mapping: Dict[str, str] = {}

    if lower_path.endswith((".xlsx", ".xls")):
        wb = load_workbook(mapping_file, data_only=True)
        ws = wb.active
        headers = [str(c.value) if c.value is not None else "" for c in ws[1]]
        if key_col and key_col not in headers:
            raise KeyError(f"Columna clave '{key_col}' no encontrada en archivo de mapeo: {headers}")
        if val_col and val_col not in headers:
            raise KeyError(f"Columna valor '{val_col}' no encontrada en archivo de mapeo: {headers}")
        key_idx = (headers.index(key_col) + 1) if key_col else 1
        val_idx = (headers.index(val_col) + 1) if val_col else 2
        for r in range(2, ws.max_row + 1):
            k = ws.cell(row=r, column=key_idx).value
            v = ws.cell(row=r, column=val_idx).value
            if k is None or str(k) == "":
                continue
            mapping[str(k)] = "" if v is None else str(v)
        return mapping

    if lower_path.endswith(".csv"):
        with open(mapping_file, "r", newline="", encoding="utf-8") as f:
            reader = csv.reader(f)
            try:
                headers = next(reader)
            except StopIteration:
                return {}
            headers = ["" if h is None else str(h) for h in headers]
            key_idx = headers.index(key_col) if key_col else 0
            val_idx = headers.index(val_col) if val_col else 1
            for row in reader:
                if not row:
                    continue
                k = row[key_idx] if key_idx < len(row) else ""
                v = row[val_idx] if val_idx < len(row) else ""
                if str(k).strip() == "":
                    continue
                mapping[str(k)] = str(v)
        return mapping

    raise ValueError(
        f"Formato de entrada no soportado para archivo de mapeo '{mapping_file}'. Use .xlsx, .xls o .csv"
    )


def derive_output_path(input_path: str, output_path: Optional[str]) -> str:
    if output_path:
        return output_path
    base, _ext = os.path.splitext(input_path)
    return f"{base}__modificado.xlsx"


def main(argv: Optional[List[str]] = None) -> int:
    parser = argparse.ArgumentParser(
        description=(
            "Lee un Excel/CSV, busca items por columna y agrega una columna nueva al lado "
            "con la información indicada. Exporta un nuevo Excel (.xlsx) con los cambios."
        )
    )

    parser.add_argument("input", help="Ruta del archivo de entrada (.xlsx/.xls/.csv)")
    parser.add_argument("--item-column", required=True, help="Nombre de la columna donde buscar el item")
    parser.add_argument(
        "--new-column",
        required=True,
        help="Nombre de la columna nueva a crear o actualizar (se ubica junto al item)",
    )
    parser.add_argument(
        "--output",
        help="Ruta del archivo Excel de salida (.xlsx). Por defecto: <input>__modificado.xlsx",
    )

    mode = parser.add_mutually_exclusive_group(required=True)
    mode.add_argument(
        "--mapping-file",
        help="Archivo Excel/CSV con dos columnas: clave (item) y valor (información)",
    )
    mode.add_argument(
        "--map",
        help="Cadena de pares clave=valor separados por ';', por ejemplo: 'A001=OK;A002=Pendiente'",
    )
    mode.add_argument(
        "--interactive",
        action="store_true",
        help="Modo interactivo: ingrese 'clave=valor' por consola hasta finalizar",
    )

    parser.add_argument(
        "--mapping-key-column",
        help="Nombre de la columna 'clave' en el archivo de mapeo (por defecto: primera)",
    )
    parser.add_argument(
        "--mapping-value-column",
        help="Nombre de la columna 'valor' en el archivo de mapeo (por defecto: segunda)",
    )

    parser.add_argument(
        "--match-mode",
        choices=["exact", "contains", "regex"],
        default="exact",
        help="Cómo buscar el item: exact (igual), contains (contiene), regex (expresión regular)",
    )
    parser.add_argument(
        "--case-sensitive",
        action="store_true",
        help="Hacer la búsqueda sensible a mayúsculas/minúsculas (por defecto: no)",
    )
    parser.add_argument(
        "--no-strip",
        action="store_true",
        help="No recortar espacios en blanco al comparar (por defecto: recorta)",
    )
    parser.add_argument(
        "--skip-existing",
        action="store_true",
        help="No sobrescribir celdas ya rellenadas en la columna nueva",
    )
    parser.add_argument("--pair-sep", default=";", help="Separador entre pares en --map (por defecto: ';')")
    parser.add_argument("--kv-sep", default="=", help="Separador clave/valor en --map (por defecto: '=')")

    args = parser.parse_args(argv)

    try:
        wb, ws = load_worksheet_from_input(args.input)
    except Exception as exc:
        print(f"Error al leer '{args.input}': {exc}", file=sys.stderr)
        return 2

    # Validate item column and create/move new column next to it
    try:
        new_col_idx = ensure_new_column_position(ws, args.item_column, args.new_column)
    except Exception as exc:
        print(f"Error preparando la columna nueva: {exc}", file=sys.stderr)
        return 2

    # Re-read headers to find the (possibly shifted) item column index
    headers_after = [str(c.value) if c.value is not None else "" for c in ws[1]]
    if args.item_column not in headers_after:
        print(
            f"La columna de item '{args.item_column}' no existe. Encabezados: {headers_after}",
            file=sys.stderr,
        )
        return 2
    item_col_idx = headers_after.index(args.item_column) + 1

    strip_values = not args.no_strip
    case_sensitive = args.case_sensitive

    total_updates = 0

    if args.mapping_file:
        try:
            mapping = load_mapping_from_file(
                args.mapping_file, args.mapping_key_column, args.mapping_value_column
            )
        except Exception as exc:
            print(f"Error al leer el archivo de mapeo: {exc}", file=sys.stderr)
            return 2
        if not mapping:
            print("Archivo de mapeo vacío o sin claves válidas.", file=sys.stderr)
            return 2
        total_updates, keys_with_hits = apply_mapping_dict(
            ws,
            item_col_idx,
            new_col_idx,
            mapping,
            args.match_mode,
            case_sensitive,
            strip_values,
            args.skip_existing,
        )
        print(f"Se actualizaron {total_updates} celdas para {keys_with_hits} claves encontradas.")

    elif args.map:
        try:
            mapping = parse_map_string(args.map, args.pair_sep, args.kv_sep)
        except Exception as exc:
            print(f"Error en --map: {exc}", file=sys.stderr)
            return 2
        total_updates, keys_with_hits = apply_mapping_dict(
            ws,
            item_col_idx,
            new_col_idx,
            mapping,
            args.match_mode,
            case_sensitive,
            strip_values,
            args.skip_existing,
        )
        print(f"Se actualizaron {total_updates} celdas para {keys_with_hits} claves encontradas.")

    elif args.interactive:
        print(
            "Modo interactivo: escriba 'clave=valor' y presione Enter.\n"
            "Comandos: 'fin' para terminar, 'ayuda' para ver opciones."
        )
        while True:
            try:
                line = input("> ").strip()
            except (EOFError, KeyboardInterrupt):
                print("\nFinalizando...")
                break

            if not line:
                continue
            if line.lower() in {"fin", "exit", "salir", "q"}:
                break
            if line.lower() in {"ayuda", "help"}:
                print(
                    "Ingrese pares 'clave=valor'. Ejemplo: A001=OK\n"
                    f"item_column='{args.item_column}', new_column='{args.new_column}', "
                    f"match_mode='{args.match_mode}', case_sensitive={case_sensitive}, strip={strip_values}, skip_existing={args.skip_existing}"
                )
                continue
            if args.kv_sep not in line:
                print(f"Formato inválido. Use clave{args.kv_sep}valor")
                continue
            key, mapped_value = line.split(args.kv_sep, 1)
            hits = apply_single_mapping(
                ws,
                item_col_idx,
                new_col_idx,
                key,
                mapped_value,
                args.match_mode,
                case_sensitive,
                strip_values,
                args.skip_existing,
            )
            total_updates += hits
            print(f"Actualizadas {hits} filas para clave '{key}'. Total acumulado: {total_updates}")

    output_path = derive_output_path(args.input, args.output)
    try:
        # Always write xlsx
        os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
        wb.save(output_path)
    except Exception as exc:
        print(f"Error al escribir el archivo de salida: {exc}", file=sys.stderr)
        return 2

    print(f"Archivo generado: {output_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
# Actualizador de Excel por búsqueda de items

Este programa lee un archivo Excel/CSV, permite buscar items por una columna, y agrega una columna nueva con información justo al lado de esa columna de item. Genera un nuevo archivo Excel (.xlsx) con los cambios.

## Requisitos

- Python 3.8+
- Dependencias de Python:
  - openpyxl

Instalación de dependencias (si es necesario):

```bash
python3 -m pip install --break-system-packages openpyxl==3.1.5
```

También puedes instalar desde `requirements.txt` si tu entorno lo permite:

```bash
python3 -m pip install -r requirements.txt
```

## Uso básico

```bash
python3 excel_updater.py INPUT.xlsx --item-column "Item" --new-column "Info" --map "A001=OK;A002=Pendiente" --output SALIDA.xlsx
```

- `--item-column`: nombre de la columna donde están los items a buscar (debe existir en la primera fila del archivo).
- `--new-column`: nombre de la columna nueva (se insertará a la derecha de `--item-column`).
- `--map`: pares `clave=valor` separados por `;` (configurable con `--pair-sep` y `--kv-sep`).
- `--output`: ruta del Excel de salida (si no se indica, se crea `<INPUT>__modificado.xlsx`).

## Usando archivo de mapeo (Excel/CSV)

Si ya tienes un archivo con dos columnas (clave y valor):

```bash
python3 excel_updater.py INPUT.xlsx \
  --item-column "SKU" \
  --new-column "Estado" \
  --mapping-file MAPEO.xlsx \
  --mapping-key-column "SKU" \
  --mapping-value-column "Estado"
```

- Si no especificas `--mapping-key-column` y `--mapping-value-column`, se usan la primera y la segunda columna del archivo de mapeo, respectivamente.
- Si `MAPEO.csv`, la primera fila debe contener los nombres de las columnas.

## Modo interactivo

```bash
python3 excel_updater.py INPUT.xlsx --item-column "Producto" --new-column "Nota" --interactive
```

- Escribe líneas `clave=valor` y presiona Enter para aplicar.
- Comandos: `fin` para terminar, `ayuda` para ver opciones.

## Opciones de búsqueda

- `--match-mode`: `exact` (igual), `contains` (contiene), `regex` (expresión regular).
- `--case-sensitive`: activa coincidencia sensible a mayúsculas/minúsculas (por defecto, no sensible).
- `--no-strip`: no recortar espacios al comparar (por defecto, recorta).
- `--skip-existing`: no sobrescribir celdas ya rellenadas en la columna nueva.

## Formatos soportados

- Entrada: `.xlsx`, `.xls` o `.csv` (se asume que la primera fila es el encabezado).
- Salida: `.xlsx`.

## Notas

- El archivo de entrada no se modifica; el resultado se guarda en un nuevo Excel.
- La nueva columna se posiciona a la derecha de la columna de item indicada.
- Si usas `contains` o `regex`, una fila puede coincidir con múltiples claves; la última clave aplicada puede sobrescribir valores previos, a menos que uses `--skip-existing`.

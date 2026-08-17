from __future__ import annotations

import argparse
import re
from datetime import date, datetime
from pathlib import Path

import pandas as pd

from config import INPUT_DIR, PROCESSING_DIR

PRICE_SHEETS = (
    "Prices with taxes",
    "Prices wo taxes",
)
CONSUMER_ROW = 0
HEADER_ROW = 1
UNIT_ROW = 2
DATA_START_ROW = 3
DATE_COLUMN = 0
DECIMAL_PLACES = 4

FUEL_TYPES = (
    "Gas oil automobile",
    "Automotive gas oil",
    "Dieselkraftstoff (I)",
)

COUNTRY_CODE_PATTERN = re.compile(r"^([A-Z]{2})_")
UNIT_NUMBER_PATTERN = re.compile(r"(\d+(?:[.,]\d+)?)")


def list_fsc_files(input_dir: Path | None = None) -> list[Path]:
    """Return available .xlsx files in the input folder."""
    folder = input_dir or INPUT_DIR
    if not folder.exists():
        raise FileNotFoundError(f"Input folder not found: {folder}")

    files = sorted(
        path for path in folder.glob("*.xlsx")
        if not path.name.startswith("~$")
    )
    if not files:
        raise FileNotFoundError(f"No .xlsx files found in: {folder}")

    return files


def choose_input_path() -> tuple[Path, bool]:
    """
    Prompt the user to choose a single FSC file or process the whole input folder.

    Returns the selected path and whether folder mode is enabled.
    """
    files = list_fsc_files()

    print("Available FSC files:")
    for index, file_path in enumerate(files, start=1):
        print(f"  {index}. {file_path.name}")
    print("  f. Process all files in 'input/fsc file' folder")

    while True:
        choice = input("Enter file number or 'f' for folder: ").strip().lower()
        if choice == "f":
            return INPUT_DIR, True

        if not choice.isdigit():
            print("Please enter a valid file number or 'f'.")
            continue

        selected_index = int(choice)
        if 1 <= selected_index <= len(files):
            return files[selected_index - 1], False

        print(f"Please enter a number between 1 and {len(files)}, or 'f'.")


def round_numeric(value: float, decimals: int = DECIMAL_PLACES) -> float:
    """Round a numeric value to the configured number of decimal places."""
    return round(float(value), decimals)


def parse_target_month(raw_value: str | None = None) -> tuple[int, int]:
    """
    Parse a month/year value from user input.

    Accepts empty input (current month), MM, MM.YYYY, MM/YYYY, or YYYY-MM.
    """
    today = date.today()
    value = (raw_value if raw_value is not None else input(
        "Enter month to calculate (MM, MM.YYYY, or Enter for current month): "
    )).strip()

    if not value:
        return today.month, today.year

    normalized = value.replace("/", ".").replace("-", ".")
    parts = [part for part in normalized.split(".") if part]

    if len(parts) == 1 and parts[0].isdigit():
        month = int(parts[0])
        if not 1 <= month <= 12:
            raise ValueError("Month must be between 1 and 12.")
        return month, today.year

    if len(parts) == 2 and all(part.isdigit() for part in parts):
        left, right = (int(parts[0]), int(parts[1]))
        if left > 31:
            year, month = left, right
        else:
            month, year = left, right

        if not 1 <= month <= 12:
            raise ValueError("Month must be between 1 and 12.")
        if year < 100:
            year += 2000
        return month, year

    for fmt in ("%m.%Y", "%Y.%m", "%m-%Y", "%Y-%m"):
        try:
            parsed = datetime.strptime(value, fmt)
            return parsed.month, parsed.year
        except ValueError:
            continue

    raise ValueError(
        "Could not parse month. Use MM, MM.YYYY, YYYY-MM, or press Enter for current month."
    )


def parse_excel_date(value: object) -> date | None:
    """Convert an Excel cell value to a date."""
    if pd.isna(value):
        return None

    if isinstance(value, datetime):
        return value.date()

    if isinstance(value, date):
        return value

    if isinstance(value, (int, float)):
        parsed = pd.to_datetime(value, unit="D", origin="1899-12-30", errors="coerce")
        if pd.notna(parsed):
            return parsed.date()

    text = str(value).strip()
    if not text:
        return None

    for fmt in ("%d.%m.%y", "%d.%m.%Y", "%Y-%m-%d", "%d/%m/%Y", "%d/%m/%y"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue

    parsed = pd.to_datetime(text, dayfirst=True, errors="coerce")
    if pd.notna(parsed):
        return parsed.date()

    return None


def normalize_header_text(value: object) -> str:
    if pd.isna(value):
        return ""
    return re.sub(r"\s+", " ", str(value).strip()).casefold()


def extract_country_code(value: object) -> str | None:
    if pd.isna(value):
        return None

    text = str(value).strip().upper()
    if text.startswith("EU"):
        return None

    match = COUNTRY_CODE_PATTERN.match(text)
    if match:
        return match.group(1)

    return None


def is_country_consumer(value: object) -> bool:
    return extract_country_code(value) is not None


def matches_fuel_type(value: object) -> bool:
    header = normalize_header_text(value)
    if not header:
        return False

    return any(normalize_header_text(fuel) in header for fuel in FUEL_TYPES)


def extract_unit_divisor(value: object) -> float:
    if pd.isna(value):
        raise ValueError("Missing unit value in header row.")

    match = UNIT_NUMBER_PATTERN.search(str(value))
    if not match:
        raise ValueError(f"Could not find numeric unit in header cell: {value!r}")

    number = match.group(1).replace(",", ".")
    divisor = float(number)
    if divisor == 0:
        raise ValueError(f"Unit divisor must not be zero: {value!r}")

    return divisor


def select_relevant_columns(raw_df: pd.DataFrame) -> list[dict[str, object]]:
    """Find country/fuel columns that match the requested layout."""
    selected: list[dict[str, object]] = []

    for column_index in range(1, raw_df.shape[1]):
        consumer = raw_df.iat[CONSUMER_ROW, column_index]
        header = raw_df.iat[HEADER_ROW, column_index]
        unit = raw_df.iat[UNIT_ROW, column_index]
        country_code = extract_country_code(consumer)

        if country_code is None:
            continue
        if not matches_fuel_type(header):
            continue

        selected.append(
            {
                "column_index": column_index,
                "consumer": country_code,
                "fuel_type": str(header).strip(),
                "unit_divisor": extract_unit_divisor(unit),
            }
        )

    if not selected:
        raise ValueError(
            "No relevant country/fuel columns found. "
            f"Expected country codes like AT_... in row 1, {FUEL_TYPES} in row 2, "
            "and units like '1000 l' in row 3."
        )

    return selected


def choose_month_rows(raw_df: pd.DataFrame, month: int, year: int) -> list[int]:
    """Return up to two row indexes for the requested month."""
    month_rows: list[tuple[date, int]] = []

    for row_index in range(DATA_START_ROW, len(raw_df)):
        row_date = parse_excel_date(raw_df.iat[row_index, DATE_COLUMN])
        if row_date is None:
            continue
        if row_date.month == month and row_date.year == year:
            month_rows.append((row_date, row_index))

    if not month_rows:
        raise ValueError(f"No dates found for {month:02d}.{year}.")

    month_rows.sort(key=lambda item: item[0])
    first_date, first_row = month_rows[0]
    selected_rows = [first_row]

    if len(month_rows) == 1:
        return selected_rows

    target_day = date(year, month, 15)
    remaining = [(row_date, row_index) for row_date, row_index in month_rows if row_index != first_row]
    closest_row = min(
        remaining,
        key=lambda item: abs((item[0] - target_day).days),
    )[1]
    selected_rows.append(closest_row)

    if first_date.day == 15 and len(month_rows) == 2:
        return selected_rows

    return selected_rows


def build_result_dataframe(
    raw_df: pd.DataFrame,
    month: int,
    year: int,
    source_file: Path | None = None,
) -> pd.DataFrame:
    """Extract and normalize the relevant prices for the requested month."""
    columns = select_relevant_columns(raw_df)
    row_indexes = choose_month_rows(raw_df, month, year)

    records: list[dict[str, object]] = []
    for row_index in row_indexes:
        row_date = parse_excel_date(raw_df.iat[row_index, DATE_COLUMN])
        if row_date is None:
            continue

        for column in columns:
            raw_value = raw_df.iat[row_index, column["column_index"]]
            if pd.isna(raw_value):
                continue

            normalized_price = round_numeric(
                float(raw_value) / float(column["unit_divisor"])
            )
            record: dict[str, object] = {
                "date": row_date,
                "consumer": column["consumer"],
                "fuel_type": column["fuel_type"],
                "raw_value": round_numeric(raw_value),
                "unit_divisor": round_numeric(column["unit_divisor"]),
                "price_per_unit": normalized_price,
            }
            if source_file is not None:
                record["source_file"] = source_file.name
            records.append(record)

    if not records:
        raise ValueError(f"No price values found for {month:02d}.{year}.")

    result = pd.DataFrame(records)
    result.sort_values(["date", "consumer", "fuel_type"], inplace=True)
    result.reset_index(drop=True, inplace=True)
    return result


def load_prices_sheet(file_path: Path, sheet_name: str) -> pd.DataFrame:
    """Load a price sheet without assuming a header row."""
    try:
        return pd.read_excel(
            file_path,
            sheet_name=sheet_name,
            header=None,
            engine="openpyxl",
        )
    except ValueError as exc:
        raise ValueError(
            f"Sheet '{sheet_name}' was not found in {file_path.name}."
        ) from exc


def save_result_dataframes(
    result_sheets: dict[str, pd.DataFrame],
    source_file: Path,
    month: int,
    year: int,
    output_dir: Path | None = None,
) -> Path:
    """Save processed result DataFrames to the processing folder."""
    folder = output_dir or PROCESSING_DIR
    folder.mkdir(parents=True, exist_ok=True)

    output_path = folder / f"{source_file.stem}_prices_{month:02d}_{year}.xlsx"
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for sheet_name, result_df in result_sheets.items():
            result_df.to_excel(writer, sheet_name=sheet_name, index=False)

    return output_path


def process_fsc_file(
    file_path: Path,
    month: int,
    year: int,
    output_dir: Path | None = None,
    include_source_column: bool = False,
) -> tuple[dict[str, pd.DataFrame], Path]:
    """Process one FSC file and save both price tabs as xlsx."""
    result_sheets: dict[str, pd.DataFrame] = {}

    for sheet_name in PRICE_SHEETS:
        raw_df = load_prices_sheet(file_path, sheet_name)
        result_sheets[sheet_name] = build_result_dataframe(
            raw_df,
            month=month,
            year=year,
            source_file=file_path if include_source_column else None,
        )

    output_path = save_result_dataframes(
        result_sheets,
        source_file=file_path,
        month=month,
        year=year,
        output_dir=output_dir,
    )
    return result_sheets, output_path


def run_processor(
    input_path: Path | None = None,
    folder_mode: bool | None = None,
    month: int | None = None,
    year: int | None = None,
    output_dir: Path | None = None,
) -> list[tuple[Path, dict[str, pd.DataFrame], Path]]:
    """Run interactive or scripted FSC price processing."""
    if input_path is None or folder_mode is None:
        selected_path, selected_folder_mode = choose_input_path()
    else:
        selected_path, selected_folder_mode = input_path, folder_mode

    if month is None or year is None:
        month, year = parse_target_month()

    files = sorted(selected_path.glob("*.xlsx")) if selected_folder_mode else [selected_path]
    if not files:
        raise FileNotFoundError(f"No .xlsx files found in: {selected_path}")

    results: list[tuple[Path, dict[str, pd.DataFrame], Path]] = []
    for file_path in files:
        if file_path.name.startswith("~$"):
            continue

        print(f"\nProcessing: {file_path.name}")
        result_sheets, output_path = process_fsc_file(
            file_path=file_path,
            month=month,
            year=year,
            output_dir=output_dir,
            include_source_column=selected_folder_mode,
        )
        results.append((file_path, result_sheets, output_path))
        for sheet_name, result_df in result_sheets.items():
            print(f"  {sheet_name}: {len(result_df)} rows")
        print(f"  Saved to: {output_path}")

    return results


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Extract normalized FSC fuel prices from price sheets.",
    )
    parser.add_argument(
        "--input",
        type=Path,
        help="Path to one FSC xlsx file or an input folder.",
    )
    parser.add_argument(
        "--folder",
        action="store_true",
        help="Treat --input as a folder and process all xlsx files inside it.",
    )
    parser.add_argument(
        "--month",
        help="Target month (MM, MM.YYYY, YYYY-MM). Defaults to current month.",
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        help="Folder where result xlsx files will be saved (default: processing).",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()

    if args.input is None:
        run_processor(output_dir=args.output_dir)
        return

    if args.month:
        month, year = parse_target_month(args.month)
    else:
        today = date.today()
        month, year = today.month, today.year

    input_path = args.input
    folder_mode = args.folder or input_path.is_dir()
    if not folder_mode and not input_path.is_file():
        raise FileNotFoundError(f"Input file not found: {input_path}")

    run_processor(
        input_path=input_path,
        folder_mode=folder_mode,
        month=month,
        year=year,
        output_dir=args.output_dir,
    )


if __name__ == "__main__":
    main()

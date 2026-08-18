from __future__ import annotations

import argparse
from dataclasses import dataclass
from datetime import date
from pathlib import Path

from config import INPUT_DIR, setup_paths

setup_paths()

from fsc_calculator import (
    calculate_fsc_from_processed_file,
    find_calculation_basis_file,
)
from fsc_prices_processor import (
    INPUT_DIR,
    choose_input_path,
    parse_target_month,
    process_fsc_file,
)
from rate_card_updater import find_rate_card_export, update_rate_card_export


@dataclass
class FilePipelineResult:
    fsc_input: Path
    processed_file: Path
    fsc_output: Path
    rate_card_output: Path
    price_rows: dict[str, int]
    fsc_rows: dict[str, int]
    rate_card_updates: int


@dataclass
class PipelineResult:
    month: int
    year: int
    basis_file: Path
    files: list[FilePipelineResult]


def run_pipeline(
    fsc_path: Path | None = None,
    folder_mode: bool | None = None,
    month: int | None = None,
    year: int | None = None,
    basis_path: Path | None = None,
    rate_card_path: Path | None = None,
    processing_dir: Path | None = None,
    output_dir: Path | None = None,
) -> PipelineResult:
    """
    Run the full FSC Matrix pipeline end to end.

    1. Extract normalized diesel prices from input/fsc file -> processing
    2. Calculate FSC values using Calculation basis -> output
    3. Update Rate card export % over cost values -> output
    """
    if fsc_path is None or folder_mode is None:
        selected_path, selected_folder_mode = choose_input_path()
    else:
        selected_path, selected_folder_mode = fsc_path, folder_mode

    if month is None or year is None:
        month, year = parse_target_month()

    basis_file = basis_path or find_calculation_basis_file()
    rate_card_file = rate_card_path or find_rate_card_export()

    if selected_folder_mode:
        fsc_files = sorted(
            path for path in selected_path.glob("*.xlsx")
            if not path.name.startswith("~$")
        )
    else:
        fsc_files = [selected_path]

    if not fsc_files:
        raise FileNotFoundError(f"No .xlsx files found in: {selected_path}")

    print("=" * 60)
    print("FSC MATRIX PIPELINE")
    print("=" * 60)
    print(f"Target month:         {month:02d}.{year}")
    print(f"Calculation basis:    {basis_file.name}")
    print(f"Rate card source:     {rate_card_file.name}")
    print(f"Files to process:     {len(fsc_files)}")

    file_results: list[FilePipelineResult] = []

    for index, fsc_input in enumerate(fsc_files, start=1):
        print("\n" + "=" * 60)
        print(f"FILE {index}/{len(fsc_files)}: {fsc_input.name}")
        print("=" * 60)

        print("\nSTEP 1/3: Extract diesel prices from FSC file")
        print("-" * 60)
        price_sheets, processed_file = process_fsc_file(
            file_path=fsc_input,
            month=month,
            year=year,
            output_dir=processing_dir,
            include_source_column=selected_folder_mode,
        )
        for sheet_name, result_df in price_sheets.items():
            print(f"  {sheet_name}: {len(result_df)} rows")
        print(f"Saved to: {processed_file}")

        print("\nSTEP 2/3: Calculate FSC values")
        print("-" * 60)
        fsc_sheets, fsc_output = calculate_fsc_from_processed_file(
            processed_path=processed_file,
            basis_path=basis_file,
            output_dir=output_dir,
        )
        for sheet_name, result_df in fsc_sheets.items():
            print(f"  {sheet_name}: {len(result_df)} rows")
        print(f"Saved to: {fsc_output}")

        print("\nSTEP 3/3: Update Rate card export")
        print("-" * 60)
        rate_card_output, updated_rows, updates_by_sheet = update_rate_card_export(
            fsc_output_path=fsc_output,
            rate_card_path=rate_card_file,
            fsc_input_path=fsc_input,
            output_dir=output_dir,
        )
        print(f"Updated rows: {updated_rows}")
        print(f"Saved to: {rate_card_output}")

        file_results.append(
            FilePipelineResult(
                fsc_input=fsc_input,
                processed_file=processed_file,
                fsc_output=fsc_output,
                rate_card_output=rate_card_output,
                price_rows={name: len(df) for name, df in price_sheets.items()},
                fsc_rows={name: len(df) for name, df in fsc_sheets.items()},
                rate_card_updates=updated_rows,
            )
        )

    print("\n" + "=" * 60)
    print("PIPELINE COMPLETE")
    print("=" * 60)
    for item in file_results:
        print(f"FSC input:      {item.fsc_input}")
        print(f"Processing:     {item.processed_file}")
        print(f"FSC output:     {item.fsc_output}")
        print(f"Rate card:      {item.rate_card_output}")
        print()

    return PipelineResult(
        month=month,
        year=year,
        basis_file=basis_file,
        files=file_results,
    )


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Run the Daikin FSC Matrix pipeline end to end.",
    )
    parser.add_argument(
        "--input",
        type=Path,
        help="Path to one FSC xlsx in input/fsc file (interactive if omitted).",
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
        "--basis",
        type=Path,
        help="Path to Calculation basis.xlsx (auto-detected if omitted).",
    )
    parser.add_argument(
        "--rate-card",
        type=Path,
        help="Path to the original Rate card export.xlsx file.",
    )
    parser.add_argument(
        "--processing-dir",
        type=Path,
        help="Folder where intermediate price files will be saved.",
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        help="Folder where final FSC files will be saved.",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()

    if args.input is None:
        run_pipeline(
            basis_path=args.basis,
            rate_card_path=args.rate_card,
            processing_dir=args.processing_dir,
            output_dir=args.output_dir,
        )
        return

    if args.month:
        month, year = parse_target_month(args.month)
    else:
        today = date.today()
        month, year = today.month, today.year

    input_path = args.input
    folder_mode = args.folder or input_path == INPUT_DIR or input_path.is_dir()
    if not folder_mode and not input_path.is_file():
        raise FileNotFoundError(f"Input file not found: {input_path}")

    run_pipeline(
        fsc_path=input_path,
        folder_mode=folder_mode,
        month=month,
        year=year,
        basis_path=args.basis,
        rate_card_path=args.rate_card,
        processing_dir=args.processing_dir,
        output_dir=args.output_dir,
    )


if __name__ == "__main__":
    main()

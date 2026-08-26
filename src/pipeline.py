"""Command-line entry point for the full Fleet Forecast ETL pipeline."""

from __future__ import annotations

import argparse
import shutil
from datetime import datetime
from pathlib import Path

from src.extractors.pdf_extractor import PDFExtractor
from src.extractors.word_extractor import WordExtractor
from src.loaders.excel_loader import ExcelLoader, TestTrainInput, TestTrainValidationError
from src.models.workbook_reader import WorkbookReader
from src.reporting.quality_report import (
    add_execution_issue,
    build_quality_report,
    write_quality_report,
)
from src.transformers.transformation_layer import TransformationLayer
from src.utils.execution_lock import exclusive_run_lock
from src.utils.logging_config import configure_logging
from src.validators.validation_layer import ValidationLayer


def parse_program_date(value: str):
    try:
        return datetime.strptime(value, "%d/%m/%Y").date()
    except ValueError as error:
        raise argparse.ArgumentTypeError("Use the date format dd/mm/yyyy.") from error


def run(
    pdf_path: Path,
    word_path: Path,
    workbook_path: Path,
    output_path: Path,
    program_date=None,
    test_trains: list[TestTrainInput] | None = None,
    report_path: Path | None = None,
    log_dir: Path | None = None,
    lock_dir: Path | None = None,
    archive_dir: Path | None = None,
):
    log_dir = log_dir or output_path.parent / "logs"
    lock_dir = lock_dir or output_path.parent
    report_path = report_path or output_path.with_suffix(".quality_report.json")
    logger = configure_logging(log_dir)

    with exclusive_run_lock(lock_dir):
        logger.info("ETL run started for output=%s", output_path.name)
        pdf_result = PDFExtractor().extract(str(pdf_path))
        word_result = WordExtractor().extract(str(word_path))
        transformed = TransformationLayer().transform(
            pdf_result.commercial_services, pdf_result.reserve, word_result.operations
        )
        validated = ValidationLayer().validate(transformed, WorkbookReader().read(str(workbook_path)))
        report = build_quality_report(validated, pdf_path, word_path, workbook_path, program_date)
        write_quality_report(report, report_path)
        logger.info("Quality report written: %s", report_path.name)
        if validated.issues:
            logger.warning("Validation failed with %s issue(s)", len(validated.issues))
            raise RuntimeError(f"No output was generated. Review the quality report: {report_path}")

        try:
            result = ExcelLoader().load(
                validated,
                workbook_path,
                output_path,
                program_date=program_date,
                test_trains=test_trains,
            )
        except TestTrainValidationError as error:
            add_execution_issue(
                report,
                rule_id="UI-001",
                source="Interfaz - Pruebas",
                record_id="Captura de tren de prueba",
                description=str(error),
            )
            write_quality_report(report, report_path)
            logger.warning("Test-train input rejected: %s", error)
            raise
        if archive_dir:
            archive_run(archive_dir, pdf_path, word_path, workbook_path, output_path, report_path)
        logger.info("ETL run finished successfully: %s", output_path.name)
        return result


def archive_run(
    archive_dir: Path,
    pdf_path: Path,
    word_path: Path,
    workbook_path: Path,
    output_path: Path,
    report_path: Path,
) -> Path:
    """Store one immutable copy of the source files, output and report per run."""

    run_dir = archive_dir / datetime.now().strftime("%Y%m%d_%H%M%S")
    run_dir.mkdir(parents=True, exist_ok=False)
    for item in (pdf_path, word_path, workbook_path, output_path, report_path):
        shutil.copy2(item, run_dir / item.name)
    return run_dir


def main() -> None:
    parser = argparse.ArgumentParser(description="Generate the daily Programa workbook.")
    parser.add_argument("--pdf", type=Path, required=True, help="Fleet forecast PDF")
    parser.add_argument("--word", type=Path, required=True, help="Operations report DOCX")
    parser.add_argument("--programa", type=Path, required=True, help="Programa template XLSX")
    parser.add_argument("--output", type=Path, required=True, help="Destination XLSX")
    parser.add_argument("--date", type=parse_program_date, help="Program date: dd/mm/yyyy")
    parser.add_argument("--test-train", help="Test train number")
    parser.add_argument("--test-pks", help="Test train P.K.'s")
    parser.add_argument("--test-mr", help="Test train MR registration")
    parser.add_argument("--test-start", help="Test start time (UTC-6), HH:MM")
    parser.add_argument("--test-end", help="Test end time (UTC-6), HH:MM")
    parser.add_argument("--report", type=Path, help="Quality report JSON path")
    parser.add_argument("--log-dir", type=Path, help="Log folder")
    parser.add_argument("--lock-dir", type=Path, help="Shared-folder lock location")
    parser.add_argument("--archive-dir", type=Path, help="Archive folder for a successful run")
    args = parser.parse_args()

    program_date = args.date
    if program_date is None:
        captured_date = input("Fecha del Programa (dd/mm/aaaa, Enter para conservar B5): ").strip()
        if captured_date:
            program_date = parse_program_date(captured_date)

    test_values = (args.test_train, args.test_pks, args.test_mr, args.test_start, args.test_end)
    if any(test_values) and not all(test_values):
        parser.error("All --test-* values must be provided together.")
    if not any(test_values):
        has_test = input("Hay tren de pruebas? (s/n): ").strip().casefold()
        if has_test in {"s", "si"}:
            test_values = (
                input("Tren de pruebas: ").strip(),
                input("P.K.'s: ").strip(),
                input("MR/Matricula: ").strip(),
                input("Hora inicio UTC-6 (HH:MM): ").strip(),
                input("Hora final UTC-6 (HH:MM): ").strip(),
            )
        elif has_test not in {"n", "no", ""}:
            parser.error("Respond with s or n for the test train question.")
    test_trains = [TestTrainInput(*test_values)] if all(test_values) else None

    result = run(
        pdf_path=args.pdf,
        word_path=args.word,
        workbook_path=args.programa,
        output_path=args.output,
        program_date=program_date,
        test_trains=test_trains,
        report_path=args.report,
        log_dir=args.log_dir,
        lock_dir=args.lock_dir,
        archive_dir=args.archive_dir,
    )
    print(f"Programa generated: {result.output_path}")


if __name__ == "__main__":
    main()

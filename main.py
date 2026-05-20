#!/usr/bin/env python3
"""CSV to XLSX Converter.

Handles large (500k+ rows, 670MB+), RFC 4180 compliant CSV files that
break in Excel / Google Sheets. All cells are treated as text so UUIDs,
phone numbers, leading-zero zip codes, and ranges like "$20,000 to $44,999"
are preserved exactly.

Architecture:
  - PyQt6 UI on the main thread (drag/drop, preview, progress).
  - Polars on a worker thread for parsing; pandas chunked fallback if
    Polars fails or isn't available on this platform.
  - xlsxwriter in constant_memory mode for the write side, so we never
    hold an in-memory copy of the workbook.
"""

import os
import subprocess
import sys
import traceback
from pathlib import Path
from typing import Optional

from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtGui import QDragEnterEvent, QDropEvent
from PyQt6.QtWidgets import (
    QApplication,
    QFileDialog,
    QFrame,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QMainWindow,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QSizePolicy,
    QTableWidget,
    QTableWidgetItem,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

EXCEL_MAX_ROWS = 1_048_576  # Excel's hard per-sheet limit.

# Polars is the fast path; pandas is the safety net. We probe both at
# import time so the worker thread can decide per-conversion without
# paying for it again.
try:
    import polars as pl

    POLARS_AVAILABLE = True
except Exception:
    POLARS_AVAILABLE = False

import pandas as pd
import xlsxwriter


def human_size(n: float) -> str:
    for unit in ("B", "KB", "MB", "GB"):
        if n < 1024:
            return f"{n:.1f} {unit}"
        n /= 1024
    return f"{n:.1f} TB"


# ---------- workers ---------------------------------------------------


class RowCountWorker(QThread):
    """Count rows and pull a 5-row preview without slurping the file."""

    done = pyqtSignal(int, list, list)  # count, columns, preview rows
    failed = pyqtSignal(str)

    def __init__(self, path: str):
        super().__init__()
        self.path = path

    def run(self) -> None:
        try:
            if POLARS_AVAILABLE:
                lf = pl.scan_csv(self.path, infer_schema_length=0)
                # pl.len() over a lazy frame counts rows without materializing data.
                count = lf.select(pl.len()).collect().item()
                preview = lf.head(5).collect()
                cols = list(preview.columns)
                rows = [list(r) for r in preview.iter_rows()]
                self.done.emit(int(count), cols, rows)
                return

            # Pandas fallback: preview with read_csv(nrows=5), then a cheap
            # newline count. Note this isn't exact for CSVs with quoted
            # newlines, but the progress bar uses it only as an estimate.
            preview_df = pd.read_csv(
                self.path,
                dtype=str,
                nrows=5,
                na_filter=False,
                keep_default_na=False,
            )
            cols = list(preview_df.columns)
            rows = preview_df.values.tolist()
            count = 0
            with open(self.path, "rb") as f:
                for _ in f:
                    count += 1
            self.done.emit(max(0, count - 1), cols, rows)
        except Exception as e:
            self.failed.emit(str(e))


class ConvertWorker(QThread):
    progress = pyqtSignal(int, int)  # processed, total
    status = pyqtSignal(str)
    warning = pyqtSignal(str)
    failed = pyqtSignal(str)
    done = pyqtSignal(str)  # output path

    def __init__(self, input_path: str, output_path: str, total_rows: int):
        super().__init__()
        self.input_path = input_path
        self.output_path = output_path
        self.total_rows = max(total_rows, 1)

    def run(self) -> None:
        try:
            # Try Polars first because:
            #   - infer_schema_length=0 forces every column to stay Utf8,
            #     so UUIDs, phone numbers, zips with leading zeros, and
            #     ranges like "$20,000 to $44,999" never get auto-typed.
            #   - it handles RFC 4180 quoted fields with embedded commas,
            #     quotes, and newlines correctly out of the box.
            # If Polars fails for any reason (rare schema bug, encoding
            # quirk, missing wheel on this platform), drop into the pandas
            # chunked path. It's slower but battle-tested.
            if POLARS_AVAILABLE:
                try:
                    self._convert_with_polars()
                    return
                except MemoryError:
                    raise
                except Exception as e:
                    self.warning.emit(
                        f"Polars path failed ({e}). Falling back to pandas."
                    )
            self._convert_with_pandas()
        except MemoryError:
            self.failed.emit(
                "Out of memory while converting. Close other apps and retry, "
                "or split the CSV into smaller files first."
            )
        except Exception as e:
            tb = traceback.format_exc(limit=3)
            self.failed.emit(f"{e}\n\n{tb}")

    def _convert_with_polars(self) -> None:
        self.status.emit("Reading CSV with Polars (all columns as text)...")
        df = pl.read_csv(
            self.input_path,
            infer_schema_length=0,
            ignore_errors=False,
            truncate_ragged_lines=False,
            try_parse_dates=False,
            null_values=[],  # don't treat any literal as null; keep empty as empty
        )
        self.status.emit(
            f"Loaded {df.height:,} rows x {df.width} columns. Writing XLSX..."
        )
        # Update total now that we know it exactly.
        self.total_rows = max(df.height, 1)
        self.progress.emit(0, self.total_rows)
        self._write_xlsx(df.iter_rows(), list(df.columns), df.height)

    def _convert_with_pandas(self) -> None:
        self.status.emit("Reading CSV with pandas (chunked)...")
        chunk_iter = pd.read_csv(
            self.input_path,
            dtype=str,
            na_filter=False,
            keep_default_na=False,
            chunksize=50_000,
            engine="c",
            on_bad_lines="warn",
        )
        first = next(chunk_iter)
        cols = list(first.columns)

        def row_gen():
            for row in first.itertuples(index=False, name=None):
                yield row
            for chunk in chunk_iter:
                for row in chunk.itertuples(index=False, name=None):
                    yield row

        self._write_xlsx(row_gen(), cols, self.total_rows)

    def _write_xlsx(self, row_iter, columns: list, total_rows: int) -> None:
        # constant_memory: only the current row lives in RAM, prior rows
        # are flushed to disk. Required for 500k+ row workbooks.
        workbook = xlsxwriter.Workbook(
            self.output_path,
            {
                "constant_memory": True,
                "strings_to_formulas": False,
                "strings_to_urls": False,
                "strings_to_numbers": False,
            },
        )
        header_fmt = workbook.add_format({"bold": True})

        sheet_idx = 1
        sheet = workbook.add_worksheet(f"Data_{sheet_idx}")
        self._write_header(sheet, columns, header_fmt)

        max_data_rows = EXCEL_MAX_ROWS - 1  # leave row 0 for the header
        row_in_sheet = 1
        processed = 0
        emit_every = max(1, min(5000, total_rows // 100 or 1))

        for row in row_iter:
            if row_in_sheet > max_data_rows:
                sheet_idx += 1
                sheet = workbook.add_worksheet(f"Data_{sheet_idx}")
                self._write_header(sheet, columns, header_fmt)
                row_in_sheet = 1
                self.warning.emit(
                    f"Per-sheet row limit reached, continuing in sheet Data_{sheet_idx}."
                )

            for c, val in enumerate(row):
                # None / "" → leave the cell empty rather than writing
                # the strings "NaN" or "None".
                if val is None or val == "":
                    continue
                sheet.write_string(row_in_sheet, c, str(val))

            row_in_sheet += 1
            processed += 1

            if processed % emit_every == 0:
                self.progress.emit(processed, max(total_rows, processed))
                self.status.emit(
                    f"Wrote {processed:,} / {max(total_rows, processed):,} rows..."
                )

        workbook.close()
        self.progress.emit(processed, max(total_rows, processed))
        self.status.emit(f"Done. Wrote {processed:,} rows to {self.output_path}")
        self.done.emit(self.output_path)

    @staticmethod
    def _write_header(sheet, columns: list, fmt) -> None:
        for c, name in enumerate(columns):
            sheet.write_string(0, c, "" if name is None else str(name), fmt)


# ---------- widgets ---------------------------------------------------


class DropZone(QFrame):
    file_dropped = pyqtSignal(str)
    clicked = pyqtSignal()

    def __init__(self):
        super().__init__()
        self.setObjectName("DropZone")
        self.setAcceptDrops(True)
        self.setMinimumHeight(160)
        self.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred
        )
        self.setCursor(Qt.CursorShape.PointingHandCursor)

        layout = QVBoxLayout(self)
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title = QLabel("Drop a .csv file here")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        f = title.font()
        f.setPointSize(16)
        f.setBold(True)
        title.setFont(f)
        sub = QLabel("or use the Choose File button below")
        sub.setAlignment(Qt.AlignmentFlag.AlignCenter)
        sub.setStyleSheet("color: #9d9d9d;")
        layout.addWidget(title)
        layout.addWidget(sub)

        self._active = False
        self._restyle()

    def _restyle(self) -> None:
        border = "#007acc" if self._active else "#3e3e42"
        bg = "#2d2d30" if self._active else "#252526"
        self.setStyleSheet(
            f"QFrame#DropZone {{ background-color: {bg}; "
            f"border: 2px dashed {border}; border-radius: 8px; }}"
        )

    def dragEnterEvent(self, event: QDragEnterEvent) -> None:
        urls = event.mimeData().urls() if event.mimeData().hasUrls() else []
        if any(u.toLocalFile().lower().endswith(".csv") for u in urls):
            event.acceptProposedAction()
            self._active = True
            self._restyle()

    def dragLeaveEvent(self, event) -> None:
        self._active = False
        self._restyle()

    def dropEvent(self, event: QDropEvent) -> None:
        for url in event.mimeData().urls():
            path = url.toLocalFile()
            if path.lower().endswith(".csv"):
                self.file_dropped.emit(path)
                break
        self._active = False
        self._restyle()

    def mousePressEvent(self, event) -> None:
        if event.button() == Qt.MouseButton.LeftButton:
            self.clicked.emit()


# ---------- main window -----------------------------------------------


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("CSV to XLSX Converter")
        self.resize(920, 800)

        self.input_path: Optional[str] = None
        self.output_path: Optional[str] = None
        self.row_count: int = 0
        self.warnings_log: list[str] = []
        self.count_worker: Optional[RowCountWorker] = None
        self.convert_worker: Optional[ConvertWorker] = None

        root = QWidget()
        self.setCentralWidget(root)
        layout = QVBoxLayout(root)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(12)

        # Drop zone
        self.drop_zone = DropZone()
        self.drop_zone.file_dropped.connect(self.set_input_file)
        self.drop_zone.clicked.connect(self.choose_input_file)
        layout.addWidget(self.drop_zone)

        # File row
        file_row = QHBoxLayout()
        self.choose_btn = QPushButton("Choose File…")
        self.choose_btn.clicked.connect(self.choose_input_file)
        self.file_info = QLabel("No file selected")
        self.file_info.setStyleSheet("color: #9d9d9d;")
        file_row.addWidget(self.choose_btn)
        file_row.addWidget(self.file_info, 1)
        layout.addLayout(file_row)

        # Preview
        layout.addWidget(QLabel("Preview (first 5 rows):"))
        self.preview = QTableWidget(0, 0)
        self.preview.setMaximumHeight(170)
        self.preview.horizontalHeader().setSectionResizeMode(
            QHeaderView.ResizeMode.Interactive
        )
        self.preview.verticalHeader().setVisible(False)
        layout.addWidget(self.preview)

        # Output row
        out_row = QHBoxLayout()
        self.choose_out_btn = QPushButton("Choose Output Location…")
        self.choose_out_btn.clicked.connect(self.choose_output_path)
        self.out_label = QLabel("Output: (defaults to same folder as input)")
        self.out_label.setStyleSheet("color: #9d9d9d;")
        out_row.addWidget(self.choose_out_btn)
        out_row.addWidget(self.out_label, 1)
        layout.addLayout(out_row)

        # Convert
        self.convert_btn = QPushButton("Convert")
        self.convert_btn.setEnabled(False)
        self.convert_btn.clicked.connect(self.start_convert)
        f = self.convert_btn.font()
        f.setBold(True)
        f.setPointSize(f.pointSize() + 1)
        self.convert_btn.setFont(f)
        layout.addWidget(self.convert_btn)

        # Progress + status
        self.progress = QProgressBar()
        self.progress.setValue(0)
        layout.addWidget(self.progress)
        self.status = QLabel("Idle.")
        self.status.setWordWrap(True)
        layout.addWidget(self.status)

        # Warnings (collapsible)
        self.warnings_group = QGroupBox("Warnings (0)")
        self.warnings_group.setCheckable(True)
        self.warnings_group.setChecked(False)
        wlay = QVBoxLayout(self.warnings_group)
        self.warnings_view = QTextEdit()
        self.warnings_view.setReadOnly(True)
        self.warnings_view.setMaximumHeight(140)
        wlay.addWidget(self.warnings_view)
        self.warnings_group.toggled.connect(self.warnings_view.setVisible)
        self.warnings_view.setVisible(False)
        layout.addWidget(self.warnings_group)

        # Open folder (hidden until conversion succeeds)
        self.open_folder_btn = QPushButton("Open Output Folder")
        self.open_folder_btn.clicked.connect(self.open_output_folder)
        self.open_folder_btn.setVisible(False)
        layout.addWidget(self.open_folder_btn)

    # ---- input ----

    def choose_input_file(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Choose CSV",
            str(Path.home()),
            "CSV files (*.csv);;All files (*)",
        )
        if path:
            self.set_input_file(path)

    def set_input_file(self, path: str) -> None:
        if not os.path.isfile(path):
            QMessageBox.warning(self, "Not a file", f"Can't read: {path}")
            return
        self.input_path = path
        size = os.path.getsize(path)
        self.file_info.setText(
            f"{Path(path).name} — {human_size(size)} — counting rows…"
        )
        self.set_default_output(path)
        self._refresh_convert_enabled()

        if self.count_worker is not None and self.count_worker.isRunning():
            self.count_worker.requestInterruption()
        self.count_worker = RowCountWorker(path)
        self.count_worker.done.connect(self._on_count_done)
        self.count_worker.failed.connect(self._on_count_failed)
        self.count_worker.start()

    def _on_count_done(self, count: int, cols: list, rows: list) -> None:
        if not self.input_path:
            return
        self.row_count = count
        size = os.path.getsize(self.input_path)
        self.file_info.setText(
            f"{Path(self.input_path).name} — {human_size(size)} — {count:,} rows"
        )
        self._populate_preview(cols, rows)
        self._refresh_convert_enabled()

    def _on_count_failed(self, msg: str) -> None:
        self.row_count = 0
        if self.input_path:
            size = os.path.getsize(self.input_path)
            self.file_info.setText(
                f"{Path(self.input_path).name} — {human_size(size)} — row count unavailable"
            )
        self._log_warning(f"Could not count rows: {msg}")

    def _populate_preview(self, cols: list, rows: list) -> None:
        self.preview.clear()
        self.preview.setColumnCount(len(cols))
        self.preview.setHorizontalHeaderLabels([str(c) for c in cols])
        self.preview.setRowCount(len(rows))
        for r, row in enumerate(rows):
            for c, val in enumerate(row):
                self.preview.setItem(
                    r, c, QTableWidgetItem("" if val is None else str(val))
                )

    def set_default_output(self, input_path: str) -> None:
        self.output_path = str(Path(input_path).with_suffix(".xlsx"))
        self.out_label.setText(f"Output: {self.output_path}")

    def choose_output_path(self) -> None:
        default = self.output_path or str(Path.home() / "output.xlsx")
        path, _ = QFileDialog.getSaveFileName(
            self, "Save XLSX as", default, "Excel workbook (*.xlsx)"
        )
        if not path:
            return
        if not path.lower().endswith(".xlsx"):
            path += ".xlsx"
        if not self._output_writable(path):
            QMessageBox.warning(
                self,
                "Cannot write here",
                "That folder isn't writable. Pick another location.",
            )
            return
        self.output_path = path
        self.out_label.setText(f"Output: {self.output_path}")
        self._refresh_convert_enabled()

    @staticmethod
    def _output_writable(path: str) -> bool:
        folder = os.path.dirname(path) or "."
        return os.access(folder, os.W_OK)

    def _refresh_convert_enabled(self) -> None:
        self.convert_btn.setEnabled(bool(self.input_path and self.output_path))

    # ---- convert ----

    def start_convert(self) -> None:
        if not self.input_path or not self.output_path:
            return
        if not self._output_writable(self.output_path):
            QMessageBox.warning(
                self,
                "Output not writable",
                "Pick a different output location — that folder isn't writable.",
            )
            return

        self.warnings_log.clear()
        self.warnings_view.clear()
        self.warnings_group.setTitle("Warnings (0)")
        self.open_folder_btn.setVisible(False)
        self.convert_btn.setEnabled(False)
        self.choose_btn.setEnabled(False)
        self.choose_out_btn.setEnabled(False)

        if self.row_count > 0:
            self.progress.setRange(0, self.row_count)
            self.progress.setValue(0)
        else:
            self.progress.setRange(0, 0)  # indeterminate marquee
        self.status.setText("Starting conversion…")

        self.convert_worker = ConvertWorker(
            self.input_path, self.output_path, self.row_count
        )
        self.convert_worker.progress.connect(self._on_progress)
        self.convert_worker.status.connect(self.status.setText)
        self.convert_worker.warning.connect(self._log_warning)
        self.convert_worker.failed.connect(self._on_failed)
        self.convert_worker.done.connect(self._on_done)
        self.convert_worker.start()

    def _on_progress(self, processed: int, total: int) -> None:
        if total > 0:
            if self.progress.maximum() != total:
                self.progress.setRange(0, total)
            self.progress.setValue(min(processed, total))

    def _log_warning(self, msg: str) -> None:
        self.warnings_log.append(msg)
        self.warnings_view.append(msg)
        self.warnings_group.setTitle(f"Warnings ({len(self.warnings_log)})")

    def _on_failed(self, msg: str) -> None:
        self.convert_btn.setEnabled(True)
        self.choose_btn.setEnabled(True)
        self.choose_out_btn.setEnabled(True)
        self.progress.setRange(0, 100)
        self.progress.setValue(0)
        self.status.setText("Conversion failed.")
        QMessageBox.critical(self, "Conversion failed", msg)

    def _on_done(self, output_path: str) -> None:
        self.convert_btn.setEnabled(True)
        self.choose_btn.setEnabled(True)
        self.choose_out_btn.setEnabled(True)
        if self.progress.maximum() == 0:
            self.progress.setRange(0, 1)
            self.progress.setValue(1)
        else:
            self.progress.setValue(self.progress.maximum())
        self.status.setText(f"Done. Saved to: {output_path}")
        self.open_folder_btn.setVisible(True)

    def open_output_folder(self) -> None:
        if not self.output_path:
            return
        subprocess.run(["open", os.path.dirname(self.output_path)])


DARK_STYLESHEET = """
QMainWindow, QWidget { background-color: #1e1e1e; color: #e0e0e0; }
QLabel { color: #e0e0e0; }
QPushButton {
    background-color: #2d2d30; color: #e0e0e0;
    border: 1px solid #3e3e42; padding: 8px 14px; border-radius: 6px;
}
QPushButton:hover:!disabled { background-color: #3e3e42; border-color: #007acc; }
QPushButton:disabled { color: #6d6d6d; }
QProgressBar {
    border: 1px solid #3e3e42; border-radius: 4px; background-color: #2d2d30;
    text-align: center; color: #e0e0e0; height: 20px;
}
QProgressBar::chunk { background-color: #007acc; border-radius: 3px; }
QTextEdit {
    background-color: #1e1e1e; color: #d4d4d4; border: 1px solid #3e3e42;
    font-family: 'Menlo', monospace; font-size: 11px;
}
QTableWidget {
    background-color: #252526; color: #e0e0e0;
    border: 1px solid #3e3e42; gridline-color: #3e3e42;
}
QHeaderView::section {
    background-color: #2d2d30; color: #e0e0e0;
    border: 1px solid #3e3e42; padding: 4px;
}
QGroupBox {
    border: 1px solid #3e3e42; border-radius: 4px;
    margin-top: 10px; padding-top: 12px; color: #e0e0e0;
}
QGroupBox::title { subcontrol-origin: margin; left: 10px; padding: 0 4px; }
"""


def main() -> None:
    app = QApplication(sys.argv)
    app.setApplicationName("CSV to XLSX Converter")
    app.setStyleSheet(DARK_STYLESHEET)
    win = MainWindow()
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()

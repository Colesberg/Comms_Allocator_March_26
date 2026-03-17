# ---------------------------------------------------------------------------
# Imports
# ---------------------------------------------------------------------------
import fitz  # PyMuPDF
import os
import re
import traceback
import pandas as pd
import xlwings as xw
import xml.etree.ElementTree as ET

from PyQt5.QtWidgets import (
    QApplication, QWidget, QDialog, QLabel, QPushButton, QVBoxLayout,
    QHBoxLayout, QFileDialog, QMessageBox, QListWidget, QComboBox,
    QTextEdit, QRadioButton, QButtonGroup
)
from PyQt5.QtCore import Qt
import sys


# ===========================================================================
# 1. UTILITIES
# ===========================================================================
# ---------------------------------------------------------------------------
# 1.1 Text / Header Normalization Utilities
# ---------------------------------------------------------------------------
# ---------------------------------------------------------------------------
# 1.1.1 Utility: Normalize Text
# ---------------------------------------------------------------------------
def normalize_text(value):
    """
    Convert a value to a clean lowercase comparison string.

    Purpose:
    - handles NaN safely
    - strips leading / trailing spaces
    - lowercases text
    """
    if pd.isna(value):
        return ""

    return str(value).strip().lower()

# ---------------------------------------------------------------------------
# 1.1.2 Utility: Normalize Text Preserve Case
# ---------------------------------------------------------------------------
def normalize_text_preserve_case(value):
    """
    Convert a value to a clean text string while preserving original casing.

    Purpose:
    - useful where we want clean display text
    - removes line breaks
    - trims leading / trailing spaces
    - collapses repeated internal spaces
    """
    if pd.isna(value):
        return ""

    value = str(value)
    value = value.replace("\n", " ").replace("\r", " ")
    value = re.sub(r"\s+", " ", value).strip()

    return value

# ---------------------------------------------------------------------------
# 1.1.3 Utility: Normalize Column Headers
# ---------------------------------------------------------------------------
def normalize_column_headers(df):
    """
    Return a copy of df with cleaned and unique column headers.

    Logic:
    - clean spacing / line breaks
    - blank headers become unnamed_#
    - duplicate headers become header_2, header_3, etc.
    """
    df = df.copy()

    cleaned_columns = []
    seen = {}

    for idx, col in enumerate(df.columns, start=1):
        cleaned_col = normalize_text_preserve_case(col)

        if cleaned_col == "":
            cleaned_col = f"unnamed_{idx}"

        if cleaned_col in seen:
            seen[cleaned_col] += 1
            cleaned_col = f"{cleaned_col}_{seen[cleaned_col]}"
        else:
            seen[cleaned_col] = 1

        cleaned_columns.append(cleaned_col)

    df.columns = cleaned_columns
    return df


# ---------------------------------------------------------------------------
# 1.2 DataFrame Search / Structure Utilities
# ---------------------------------------------------------------------------
# ---------------------------------------------------------------------------
# 1.2.1 Utility: Find Row Index By First Column Value
# ---------------------------------------------------------------------------
def find_row_index_by_first_column_value(df, target_text):
    """
    Find the first row index where column 0 matches target_text
    after normalization.
    """
    try:
        target_text_norm = normalize_text(target_text)

        for idx in df.index:
            cell_value = df.iloc[idx, 0]
            if normalize_text(cell_value) == target_text_norm:
                return idx

        return None

    except Exception as e:
        print(f"🔴 Error in find_row_index_by_first_column_value: {e}")
        print(traceback.format_exc())
        return None

# ---------------------------------------------------------------------------
# 1.2.2 Utility: Set Header From Row
# ---------------------------------------------------------------------------
def set_header_from_row(df, header_row_idx):
    """
    Use the specified row as the column header row,
    then drop that row and all rows above it.
    """
    df = df.copy()

    df.columns = df.iloc[header_row_idx]
    df = df.drop(df.index[:header_row_idx + 1])
    df = df.reset_index(drop=True)

    return df

# ---------------------------------------------------------------------------
# 1.2.3 Utility: Drop Blank Rows By Column Position
# ---------------------------------------------------------------------------
def drop_blank_rows_by_column_position(df, column_position):
    """
    Drop rows where the specified column position is blank / NaN.
    """
    df = df.copy()

    if column_position >= len(df.columns):
        print(f"🟠 column_position {column_position} is outside df columns.")
        return df

    target_col = df.columns[column_position]
    df = df.dropna(subset=[target_col])

    return df

# ---------------------------------------------------------------------------
# 1.2.4 Utility: Drop Fully Blank Rows
# ---------------------------------------------------------------------------
def drop_fully_blank_rows(df):
    """
    Drop rows where every value is blank / NaN.
    """
    df = df.copy()
    df = df.dropna(how="all")
    df = df.reset_index(drop=True)
    return df

# ---------------------------------------------------------------------------
# 1.2.5 Utility: Keep Rows Where Column Not Blank
# ---------------------------------------------------------------------------
def keep_rows_where_column_not_blank(df, column_name):
    """
    Keep only rows where the target column is not blank
    after normalization.
    """
    df = df.copy()

    if column_name not in df.columns:
        print(f"🟠 Column not found for non-blank filter: {column_name}")
        return df

    df = df[
        df[column_name].apply(lambda x: normalize_text(x) != "")
    ].reset_index(drop=True)

    return df

# ---------------------------------------------------------------------------
# 1.2.6 Utility: Keep Rows Where Column Is Numeric
# ---------------------------------------------------------------------------
def keep_rows_where_column_is_numeric(df, column_name):
    """
    Keep only rows where the target column can be converted to numeric.
    """
    df = df.copy()

    if column_name not in df.columns:
        print(f"🟠 Column not found for numeric filter: {column_name}")
        return df

    df[column_name] = pd.to_numeric(df[column_name], errors="coerce")
    df = df[df[column_name].notna()].reset_index(drop=True)

    return df

# ---------------------------------------------------------------------------
# 1.2.7 Utility: Find Matching Column
# ---------------------------------------------------------------------------
def find_matching_column(columns_list, candidate_names, allow_contains=True):
    """
    Return the first real column name that matches any candidate name
    after normalization.

    Logic:
    1. exact normalized match
    2. optional relaxed contains match
    """
    try:
        normalized_map = {
            normalize_text(col): col
            for col in columns_list
        }

        # 1) Exact normalized match
        for candidate in candidate_names:
            candidate_norm = normalize_text(candidate)
            if candidate_norm in normalized_map:
                return normalized_map[candidate_norm]

        # 2) Relaxed contains match
        if allow_contains:
            for candidate in candidate_names:
                candidate_norm = normalize_text(candidate)
                for norm_col, original_col in normalized_map.items():
                    if candidate_norm and candidate_norm in norm_col:
                        return original_col

        return None

    except Exception as e:
        print(f"🔴 Error in find_matching_column: {e}")
        print(traceback.format_exc())
        return None

# ---------------------------------------------------------------------------
# 1.2.8 Utility: Coerce Series To Numeric
# ---------------------------------------------------------------------------
def coerce_series_to_numeric(series):
    """
    Convert a pandas Series to numeric safely.
    """
    if isinstance(series, pd.DataFrame):
        raise ValueError(
            "coerce_series_to_numeric expected a Series but received a DataFrame. "
            "This usually means duplicate column names are still present."
        )

    cleaned = (
        series.astype(str)
        .str.replace("R", "", regex=False)
        .str.replace(",", "", regex=False)
        .str.replace("(", "-", regex=False)
        .str.replace(")", "", regex=False)
        .str.strip()
    )

    cleaned = cleaned.replace(
        {
            "": None,
            "nan": None,
            "None": None,
        }
    )

    return pd.to_numeric(cleaned, errors="coerce")

# ---------------------------------------------------------------------------
# 1.3 File Reading Utilities
# ---------------------------------------------------------------------------
# ---------------------------------------------------------------------------
# 1.3.1 Utility: Read Raw First Sheet
# ---------------------------------------------------------------------------
def read_raw_first_sheet(file_path, header=None):
    """
    Read the first sheet of an Excel file.


    Logic:
    - use pandas directly for xlsx / xlsm style files
    - for legacy .xls files, fall back to xlwings because pandas may require
      xlrd in the local environment
    """
    print(f"   [UTILITY] Reading raw first sheet: {file_path}")


    try:
        ext = os.path.splitext(str(file_path))[1].strip().lower()


        # -------------------------------------------------------------------
        # A: Standard pandas read for modern Excel files
        # -------------------------------------------------------------------
        if ext != ".xls":
            return pd.read_excel(file_path, sheet_name=0, header=header)


        # -------------------------------------------------------------------
        # B: xlwings fallback for legacy .xls files
        # -------------------------------------------------------------------
        print("   [UTILITY] Legacy .xls detected -> using xlwings fallback")


        app = None
        wb = None


        try:
            app = xw.App(visible=False, add_book=False)
            app.display_alerts = False
            app.screen_updating = False


            wb = app.books.open(file_path, update_links=False, read_only=True)
            sh = wb.sheets[0]


            values = sh.used_range.value


            if values is None:
                print("   [UTILITY] No used range values found.")
                return pd.DataFrame()


            if not isinstance(values, list):
                values = [[values]]
            elif len(values) > 0 and not isinstance(values[0], list):
                values = [values]


            raw_df = pd.DataFrame(values)
            print(f"   [UTILITY] xlwings raw_df shape: {raw_df.shape}")


            if header is None:
                return raw_df


            if header >= len(raw_df):
                print(f"   [UTILITY] Requested header row {header} outside raw_df range.")
                return pd.DataFrame()


            output_df = raw_df.copy()
            output_df.columns = output_df.iloc[header]
            output_df = output_df.drop(output_df.index[:header + 1]).reset_index(drop=True)


            print(f"   [UTILITY] xlwings header-applied df shape: {output_df.shape}")
            return output_df


        finally:
            try:
                if wb is not None:
                    wb.close()
            except Exception:
                pass


            try:
                if app is not None:
                    app.quit()
            except Exception:
                pass


    except Exception as e:
        print(f"🔴 Error in read_raw_first_sheet: {e}")
        print(traceback.format_exc())
        return pd.DataFrame()



# ---------------------------------------------------------------------------
# 1.3.2 Utility: Read Best Matching Sheet By Required Columns
# ---------------------------------------------------------------------------
def read_best_matching_sheet_by_required_columns(
    file_path,
    required_column_candidates,
    header=0,
    exclude_sheet_names=None
):
    """
    Read all sheets in a workbook and return the sheet DataFrame that best matches
    the required column candidates.

    Purpose:
    - useful where a workbook may contain multiple sheets
    - helps choose the best source sheet automatically
    - avoids hardcoding sheet index 0 when the layout varies

    Args:
        file_path: full file path
        required_column_candidates:
            list of candidate-name lists, for example:
            [
                ["Membership Number", "External Member Number"],
                ["Commission Amount (Excluding VAT)", "Amount to be Paid"]
            ]
        header: header row passed to pd.read_excel
        exclude_sheet_names: optional list of sheet names to skip

    Returns:
        best_sheet_name, best_df
        or (None, None) if no usable sheet is found
    """
    try:
        print("------------------------------------------------------------")
        print("🟦 1.3.2 Utility: Read Best Matching Sheet By Required Columns")
        print(f"   file_path: {file_path}")
        print(f"   header: {header}")

        if exclude_sheet_names is None:
            exclude_sheet_names = []

        exclude_sheet_names_norm = [normalize_text(x) for x in exclude_sheet_names]

        excel_file = pd.ExcelFile(file_path)
        print(f"   workbook sheets: {excel_file.sheet_names}")

        best_sheet_name = None
        best_df = None
        best_match_count = -1

        for sheet_name in excel_file.sheet_names:
            if normalize_text(sheet_name) in exclude_sheet_names_norm:
                print(f"   skipping excluded sheet: {sheet_name}")
                continue

            try:
                temp_df = pd.read_excel(file_path, sheet_name=sheet_name, header=header)

                if temp_df is None or temp_df.empty:
                    print(f"   sheet empty: {sheet_name}")
                    continue

                temp_df = normalize_column_headers(temp_df)

                match_count = 0
                for candidate_group in required_column_candidates:
                    matched_col = find_matching_column(temp_df.columns, candidate_group)
                    if matched_col is not None:
                        match_count += 1

                print(f"   sheet: {sheet_name} | match_count: {match_count}")

                if match_count > best_match_count:
                    best_match_count = match_count
                    best_sheet_name = sheet_name
                    best_df = temp_df.copy()

            except Exception as sheet_error:
                print(f"   ⚠️ Could not assess sheet '{sheet_name}': {sheet_error}")

        print(f"   best_sheet_name: {best_sheet_name}")
        print(f"   best_match_count: {best_match_count}")

        if best_sheet_name is None or best_df is None:
            return None, None

        return best_sheet_name, best_df

    except Exception as e:
        print(f"🔴 Error in read_best_matching_sheet_by_required_columns: {e}")
        print(traceback.format_exc())
        return None, None

# ---------------------------------------------------------------------------
# 1.3.3 Utility: Read Raw PDF Text Lines
# ---------------------------------------------------------------------------
def read_raw_pdf_text_lines(file_path):
    """
    Read a PDF file and return a cleaned list of text lines.

    Purpose:
    - extract text page by page
    - remove blank lines
    - preserve visible casing
    - keep line order exactly as it appears in the PDF
    """
    try:
        print(f"   [UTILITY] Reading raw PDF text lines: {file_path}")

        pdf_doc = fitz.open(file_path)
        pdf_lines = []

        for page_idx in range(len(pdf_doc)):
            page = pdf_doc.load_page(page_idx)
            page_text = page.get_text("text")

            page_lines = [
                str(line).replace("\xa0", " ").strip()
                for line in page_text.splitlines()
            ]
            page_lines = [line for line in page_lines if str(line).strip() != ""]

            print(f"   [UTILITY] Page {page_idx + 1} extracted lines: {len(page_lines)}")
            pdf_lines.extend(page_lines)

        pdf_doc.close()

        print(f"   [UTILITY] Total PDF lines extracted: {len(pdf_lines)}")
        return pdf_lines

    except Exception as e:
        print(f"🔴 Error in read_raw_pdf_text_lines: {e}")
        print(traceback.format_exc())
        return []

# ---------------------------------------------------------------------------
# 1.4 Reader Registry Utilities
# ---------------------------------------------------------------------------
# ---------------------------------------------------------------------------
# 1.4.1 Utility: Reader Registry
# ---------------------------------------------------------------------------
READER_REGISTRY = {}

# ---------------------------------------------------------------------------
# 1.4.2 Utility: Register Reader
# ---------------------------------------------------------------------------
def register_reader(comm_type):
    """
    Decorator to register a reader function against a comm_type label.
    """
    def decorator(func):
        READER_REGISTRY[comm_type] = func
        return func

    return decorator

# ---------------------------------------------------------------------------
# 1.4.3 Utility: Get Reader For Comm Type
# ---------------------------------------------------------------------------
def get_reader_for_comm_type(comm_type):
    """
    Return the registered reader function for the comm_type.
    """
    reader_func = READER_REGISTRY.get(comm_type)
    print(f"   [ENGINE] get_reader_for_comm_type('{comm_type}') -> {reader_func}")
    return reader_func

# ===========================================================================
# 2. WIDGET CLASSES
# ===========================================================================
# ---------------------------------------------------------------------------
# 2.1 Run / Mode Selection Widget Classes
# ---------------------------------------------------------------------------
# ---------------------------------------------------------------------------
# 2.1.1 Widget Class: Confirm Run Type Widget
# ---------------------------------------------------------------------------
class ConfirmRunTypeWidget(QWidget):
    def __init__(self):
        super().__init__()

        self.test_official = None
        self.file_folder = None
        self.clear_table = None
        self.user_closed_without_submit = True

        self.setWindowTitle("Confirm Process Comms Run Type")
        self.setMinimumWidth(500)

        self.init_ui()
        self.update_visibility()

    def init_ui(self):
        main_layout = QVBoxLayout()

        title = QLabel("Confirm Process Comms Run Type")
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("font-size: 16px; font-weight: bold;")
        main_layout.addWidget(title)

        info = QLabel(
            "Choose whether this is a test run or an official run.\n"
            "If official, also confirm whether you are processing a single file or a folder.\n"
            "If official + file, confirm whether the output table should be cleared."
        )
        info.setWordWrap(True)
        main_layout.addWidget(info)

        # Run Type
        run_type_layout = QHBoxLayout()
        run_type_label = QLabel("Run Type:")
        self.run_type_combo = QComboBox()
        self.run_type_combo.addItems(["test", "official"])
        self.run_type_combo.currentTextChanged.connect(self.update_visibility)
        run_type_layout.addWidget(run_type_label)
        run_type_layout.addWidget(self.run_type_combo)
        main_layout.addLayout(run_type_layout)

        # File / Folder
        self.file_folder_layout = QHBoxLayout()
        self.file_folder_label = QLabel("Import Mode:")
        self.file_folder_combo = QComboBox()
        self.file_folder_combo.addItems(["file", "folder"])
        self.file_folder_combo.currentTextChanged.connect(self.update_visibility)
        self.file_folder_layout.addWidget(self.file_folder_label)
        self.file_folder_layout.addWidget(self.file_folder_combo)
        main_layout.addLayout(self.file_folder_layout)

        # Clear Table
        self.clear_table_layout = QHBoxLayout()
        self.clear_table_label = QLabel("Clear Output Table First:")
        self.clear_table_combo = QComboBox()
        self.clear_table_combo.addItems(["yes", "no"])
        self.clear_table_layout.addWidget(self.clear_table_label)
        self.clear_table_layout.addWidget(self.clear_table_combo)
        main_layout.addLayout(self.clear_table_layout)

        # Buttons
        button_layout = QHBoxLayout()

        self.btn_submit = QPushButton("Submit")
        self.btn_submit.clicked.connect(self.submit_selection)

        self.btn_cancel = QPushButton("Cancel")
        self.btn_cancel.clicked.connect(self.cancel_selection)

        button_layout.addWidget(self.btn_submit)
        button_layout.addWidget(self.btn_cancel)

        main_layout.addLayout(button_layout)
        self.setLayout(main_layout)

    def update_visibility(self):
        run_type = self.run_type_combo.currentText().strip().lower()
        import_mode = self.file_folder_combo.currentText().strip().lower()

        if run_type == "test":
            self.file_folder_label.hide()
            self.file_folder_combo.hide()

            self.clear_table_label.hide()
            self.clear_table_combo.hide()

        elif run_type == "official":
            self.file_folder_label.show()
            self.file_folder_combo.show()

            if import_mode == "file":
                self.clear_table_label.show()
                self.clear_table_combo.show()
            else:
                self.clear_table_label.hide()
                self.clear_table_combo.hide()

    def submit_selection(self):
        run_type = self.run_type_combo.currentText().strip().lower()

        if run_type == "test":
            self.test_official = "test"
            self.file_folder = "file"
            self.clear_table = None

        elif run_type == "official":
            import_mode = self.file_folder_combo.currentText().strip().lower()

            self.test_official = "official"
            self.file_folder = import_mode

            if import_mode == "file":
                self.clear_table = self.clear_table_combo.currentText().strip().lower()
            else:
                self.clear_table = "yes"

        self.user_closed_without_submit = False
        self.close()

    def cancel_selection(self):
        reply = QMessageBox.question(
            self,
            "Cancel Confirmation",
            "Are you sure you want to cancel?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )

        if reply == QMessageBox.Yes:
            self.user_closed_without_submit = True
            self.close()

    def closeEvent(self, event):
        if self.user_closed_without_submit:
            self.user_closed_without_submit = True
        event.accept()

# ---------------------------------------------------------------------------
# 2.2 File / Folder Selection Widget Classes
# ---------------------------------------------------------------------------
# ---------------------------------------------------------------------------
# 2.2.1 Widget Class: File / Folder Picker Widget
# ---------------------------------------------------------------------------
class FileFolderPickerWidget(QDialog):
    def __init__(self, file_folder="file", parent=None):
        super().__init__(parent)

        self.file_folder = str(file_folder).strip().lower()
        self.selected_path = None

        self.setWindowTitle("Select File or Folder")
        self.resize(700, 220)
        self.setModal(True)

        self.init_ui()

    def init_ui(self):
        main_layout = QVBoxLayout()

        # Title
        title_label = QLabel("Select the required input")
        title_label.setAlignment(Qt.AlignCenter)
        title_label.setStyleSheet("font-size: 16px; font-weight: bold;")
        main_layout.addWidget(title_label)

        # Mode label
        if self.file_folder == "folder":
            mode_text = "Mode selected: Folder import"
        else:
            mode_text = "Mode selected: File import"

        self.mode_label = QLabel(mode_text)
        self.mode_label.setAlignment(Qt.AlignCenter)
        self.mode_label.setStyleSheet("font-size: 13px;")
        main_layout.addWidget(self.mode_label)

        # Selected path display
        self.path_label = QLabel("No path selected yet.")
        self.path_label.setWordWrap(True)
        self.path_label.setStyleSheet(
            "border: 1px solid #999; padding: 8px; background-color: #f7f7f7;"
        )
        main_layout.addWidget(self.path_label)

        # Browse button row
        browse_layout = QHBoxLayout()

        self.browse_button = QPushButton("Browse")
        self.browse_button.clicked.connect(self.browse_path)
        browse_layout.addWidget(self.browse_button)

        main_layout.addLayout(browse_layout)

        # Continue / Exit buttons
        button_layout = QHBoxLayout()

        self.continue_button = QPushButton("Continue")
        self.continue_button.clicked.connect(self.confirm_selection)
        button_layout.addWidget(self.continue_button)

        self.exit_button = QPushButton("Exit")
        self.exit_button.clicked.connect(self.reject)
        button_layout.addWidget(self.exit_button)

        main_layout.addLayout(button_layout)

        self.setLayout(main_layout)

    def browse_path(self):
        if self.file_folder == "folder":
            folder_path = QFileDialog.getExistingDirectory(
                self,
                "Select Folder"
            )
            if folder_path:
                self.selected_path = folder_path
                self.path_label.setText(folder_path)
        else:
            file_path, _ = QFileDialog.getOpenFileName(
                self,
                "Select File",
                "",
                "Supported Files (*.xlsx *.xls *.xlsm *.pdf)"
            )
            if file_path:
                self.selected_path = file_path
                self.path_label.setText(file_path)

    def confirm_selection(self):
        if not self.selected_path:
            QMessageBox.warning(
                self,
                "No Selection",
                "Please browse and select a valid file or folder first."
            )
            return

        self.accept()

# ---------------------------------------------------------------------------
# 2.3 Comm Type Selection Widget Classes
# ---------------------------------------------------------------------------
# ---------------------------------------------------------------------------
# 2.3.1 Widget Class: Comm Type Selection Widget
# ---------------------------------------------------------------------------
class CommTypeSelectionWidget(QDialog):
    def __init__(self, detected_comm_type=None, parent=None):
        super().__init__(parent)

        self.detected_comm_type = detected_comm_type
        self.selected_comm_type = None

        self.comm_type_options = [
            "AGBus", "AGPers", "AllanGray_Unify", "AllanGray_Silk", "Ambledown",
            "Ambledown_Silk", "Bidvest", "Bidvest_Silk", "Bonitas", "Brolink",
            "Capital_Legacy", "Capital_Legacy_Silk", "Cura", "Dischealth",
            "Dischealth_Silk", "Disclife", "Disclife_Silk", "Discinsure",
            "Discinsure_Silk", "Discgap", "Discgap_Silk", "Guardrisk", "HIC",
            "Hollard_Life", "Hollard_ST", "Liberty", "Kaelo", "Momentum_Mandy",
            "Momentum_Mandy_PDF", "Momentum_MFP", "MUA", "Nedgroup",
            "Old_Mutual_Short_Term", "Old_Mutual_Life_Invest", "Sanlam", "Santam",
            "SAU", "Sirago", "Sirago_Silk", "Stanlib", "Stratum", "Stratum_Silk",
            "Turnberry", "Zestlife"
        ]


        self.setWindowTitle("Select Commission Type")
        self.resize(500, 180)
        self.setModal(True)

        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()

        label = QLabel("Please select the commission type for this file:")
        layout.addWidget(label)

        self.combo = QComboBox()
        self.combo.addItems(self.comm_type_options)

        if self.detected_comm_type in self.comm_type_options:
            self.combo.setCurrentText(self.detected_comm_type)

        layout.addWidget(self.combo)

        button_row = QHBoxLayout()

        ok_button = QPushButton("Continue")
        ok_button.clicked.connect(self.confirm_selection)
        button_row.addWidget(ok_button)

        cancel_button = QPushButton("Exit")
        cancel_button.clicked.connect(self.reject)
        button_row.addWidget(cancel_button)

        layout.addLayout(button_row)

        self.setLayout(layout)

    def confirm_selection(self):
        self.selected_comm_type = self.combo.currentText()
        self.accept()

# ---------------------------------------------------------------------------
# 2.4.1 Widget Class: Processing Error Widget
# ---------------------------------------------------------------------------
class ProcessingErrorWidget(QDialog):
    def __init__(self, file_path, error_message, parent=None):
        super().__init__(parent)
        self.file_path = file_path
        self.error_message = error_message
        self.setWindowTitle("Processing Error")
        self.resize(700, 250)
        self.setModal(True)
        self.init_ui()
        
    def init_ui(self):
        layout = QVBoxLayout()
        title = QLabel("An error occurred while processing the file.")
        title.setStyleSheet("font-size: 14px; font-weight: bold;")
        layout.addWidget(title)
        file_label = QLabel(f"File: {self.file_path}")
        file_label.setWordWrap(True)
        layout.addWidget(file_label)
        error_box = QTextEdit()
        error_box.setReadOnly(True)
        error_box.setText(str(self.error_message))
        layout.addWidget(error_box)
        ok_button = QPushButton("OK")
        ok_button.clicked.connect(self.accept)
        layout.addWidget(ok_button)
        self.setLayout(layout)

# ===========================================================================
# 3. WIDGET LAUNCHER / WRAPPER FUNCTIONS
# ===========================================================================
# ---------------------------------------------------------------------------
# 3.1.1 Function: Confirm Test / Official / File / Folder
# ---------------------------------------------------------------------------
def f3_1_1_confirm_test_file_folder():
    """
    Launch the run-type confirmation widget.

    Returns:
        test_official, file_folder, clear_table
    """
    print("------------------------------------------------------------")
    print("🟦 3.1.1 Launch Confirm Run Type Widget")

    app = QApplication.instance()
    owns_app = False

    if app is None:
        app = QApplication(sys.argv)
        owns_app = True

    widget = ConfirmRunTypeWidget()
    widget.show()
    app.exec_()

    if widget.user_closed_without_submit:
        raise Exception("User cancelled Step 1 confirmation widget.")

    test_official = widget.test_official
    file_folder = widget.file_folder
    clear_table = widget.clear_table

    print(f"   test_official: {test_official}")
    print(f"   file_folder  : {file_folder}")
    print(f"   clear_table  : {clear_table}")

    if owns_app:
        app.quit()

    return test_official, file_folder, clear_table

# ---------------------------------------------------------------------------
# 3.2.1 Function: Prompt User To Pick File Or Folder
# ---------------------------------------------------------------------------
def f3_2_1_prompt_user_to_pick_file_or_folder(file_folder="file"):
    """
    Launch the file/folder picker widget.

    Args:
        file_folder: 'file' or 'folder'

    Returns:
        selected_path or None
    """
    print("------------------------------------------------------------")
    print("🟦 3.2.1 Launch File / Folder Picker Widget")
    print(f"   requested mode: {file_folder}")

    app = QApplication.instance()
    app_created_here = False

    if app is None:
        app = QApplication(sys.argv)
        app_created_here = True

    dialog = FileFolderPickerWidget(file_folder=file_folder)
    result = dialog.exec_()

    selected_path = dialog.selected_path if result == QDialog.Accepted else None

    print(f"   selected_path: {selected_path}")

    if app_created_here:
        app.quit()

    return selected_path

# ---------------------------------------------------------------------------
# 3.3.1 Function: Prompt Comm Type Only If Needed
# ---------------------------------------------------------------------------
def f3_3_1_prompt_comm_type_only_if_needed(detected_comm_type=None):
    """
    Launch the comm type selection widget and return the selected type.
    """
    print("------------------------------------------------------------")
    print("🟦 3.3.1 Launch Comm Type Selection Widget")
    print(f"   detected_comm_type: {detected_comm_type}")

    app = QApplication.instance()
    app_created_here = False

    if app is None:
        app = QApplication([])
        app_created_here = True

    dialog = CommTypeSelectionWidget(detected_comm_type=detected_comm_type)
    result = dialog.exec_()

    selected_comm_type = dialog.selected_comm_type if result == QDialog.Accepted else None

    print(f"   selected_comm_type: {selected_comm_type}")

    if app_created_here:
        app.quit()

    return selected_comm_type

# ---------------------------------------------------------------------------
# 3.4.1 Function: Show Processing Error Widget
# ---------------------------------------------------------------------------
def f3_4_1_show_processing_error_widget(file_path, error_message):
    """
    Show a processing error dialog for the current file.
    """
    print("------------------------------------------------------------")
    print("🟦 3.4.1 Show Processing Error Widget")
    print(f"   file_path    : {file_path}")
    print(f"   error_message: {error_message}")

    app = QApplication.instance()
    app_created_here = False

    if app is None:
        app = QApplication([])
        app_created_here = True

    dialog = ProcessingErrorWidget(file_path=file_path, error_message=error_message)
    dialog.exec_()

    if app_created_here:
        app.quit()

# ===========================================================================
# 4. COMM TYPE DETECTION / ROUTING FUNCTIONS
# ===========================================================================
# ---------------------------------------------------------------------------
# 4.1 Filename Detection Functions
# ---------------------------------------------------------------------------

# ---------------------------------------------------------------------------
# 4.1.1 Function: Detect Comm Type Label From Filename
# ---------------------------------------------------------------------------
def f4_1_1_detect_comm_type_label(file_path):
    """
    Detect the commission type label from the file name.

    Returns:
        detected label such as 'AGBus', 'AGPers', 'Bidvest', etc.
        or None if no match is found.
    """
    file_name = os.path.basename(file_path).lower().strip()

    print("------------------------------------------------------------")
    print("🟦 4.1.1 Detect comm type from filename")
    print(f"   file_name: {file_name}")

    detection_map = {
        "agbus": "AGBus",
        "agpers": "AGPers",
        "allan_gray_unify": "AllanGray_Unify",
        "allan_gray_silk": "AllanGray_Silk",
        "ambledown_silk": "Ambledown_Silk",
        "ambledown": "Ambledown",
        "bidvest_silk": "Bidvest_Silk",
        "bidvest": "Bidvest",
        "bonitas": "Bonitas",
        "brolink": "Brolink",
        "capital_legacy_unify": "Capital_Legacy",
        "capital_legacy_silk": "Capital_Legacy_Silk",
        "capital_legacy": "Capital_Legacy",
        "cura": "Cura",
        "dischealth_silk": "Dischealth_Silk",
        "dischealth": "Dischealth",
        "disclife_silk": "Disclife_Silk",
        "disclife": "Disclife",
        "discinsure_silk": "Discinsure_Silk",
        "discinsure": "Discinsure",
        "discgap_silk": "Discgap_Silk",
        "discgap": "Discgap",
        "guardrisk": "Guardrisk",
        "hic": "HIC",
        "hollard_life": "Hollard_Life",
        "hollard_st": "Hollard_ST",
        "liberty": "Liberty",
        "kaelo": "Kaelo",
        "momentum_mandy_pdf": "Momentum_Mandy_PDF",
        "momentum_silkfin": "Momentum_Mandy_PDF",
        "silkfin": "Momentum_Mandy_PDF",
        "momentum_mandy": "Momentum_Mandy",
        "momentum_mfp": "Momentum_MFP",
        "mua": "MUA",
        "nedgroup": "Nedgroup",
        "old_mutual_short_term": "Old_Mutual_Short_Term",
        "old_mutual_life_invest": "Old_Mutual_Life_Invest",
        "sanlam": "Sanlam",
        "santam": "Santam",
        "sau": "SAU",
        "sirago_silk": "Sirago_Silk",
        "sirago": "Sirago",
        "stanlib": "Stanlib",
        "stratum_silk": "Stratum_Silk",
        "stratum": "Stratum",
        "turnberry": "Turnberry",
        "zestlife": "Zestlife",
    }

    for key, value in detection_map.items():
        if key in file_name:
            print(f"✅ Detected comm type: {value}")
            return value

    print("🟠 No comm type detected from filename.")
    return None

# ---------------------------------------------------------------------------
# 4.2 Reader Routing Functions
# ---------------------------------------------------------------------------

# ---------------------------------------------------------------------------
# 4.2.1 Function: Run Engineered Reader
# ---------------------------------------------------------------------------
def f4_2_1_run_engineered_reader(comm_type, file_path, comm_month, comm_tables_main_df, reader_func=None):
    print("------------------------------------------------------------")
    print("🟦 Step 4: Run engineered reader")
    print(f"   comm_type          : {comm_type}")
    print(f"   file_path          : {file_path}")
    print(f"   comm_month         : {comm_month}")
    print(f"   comm_tables_main_df shape: {comm_tables_main_df.shape if isinstance(comm_tables_main_df, pd.DataFrame) else None}")


    if reader_func is None:
        reader_func = get_reader_for_comm_type(comm_type)


    if reader_func is None:
        print(f"🔴 No reader found for comm_type: {comm_type}")
        return None


    engineered_df = reader_func(
        file_path=file_path,
        comm_month=comm_month,
        comm_tables_main_df=comm_tables_main_df
    )


    if engineered_df is None:
        print("🔴 Reader returned None.")
        return None


    if isinstance(engineered_df, pd.DataFrame):
        print(f"✅ Reader returned DataFrame with shape: {engineered_df.shape}")
    else:
        print(f"🟠 Reader returned object type: {type(engineered_df)}")


    return engineered_df


# ===========================================================================
# 5. READER FUNCTIONS
# ===========================================================================
# ---------------------------------------------------------------------------
# 5.1 Header-Row Reader Functions
# ---------------------------------------------------------------------------

# ---------------------------------------------------------------------------
# 5.1.1 Reader: AGBus
# ---------------------------------------------------------------------------
@register_reader("AGBus")
def read_agbus_df(file_path, **kwargs):
    """
    Reader for AGBus commission statement files.

    Goal:
    - read the raw AGBus sheet correctly
    - identify the true header row
    - clean the raw data
    - map the source columns into a normalized structure

    Returns normalized columns:
    - client_name
    - product_house
    - commission_month
    - contract_number
    - total_commission
    - planner
    """
    try:
        print("------------------------------------------------------------")
        print("🧾 5.1.1 Reader: read_agbus_df")
        print(f"   file_path: {file_path}")

        # -------------------------------------------------------------------
        # Step 0: Get Variables
        # -------------------------------------------------------------------
        print("🟦 Step 0: Get passed variables")
        comm_month = kwargs.get("comm_month")
        comm_tables_main_df = kwargs.get("comm_tables_main_df")

        print(f"   comm_month: {comm_month}")
        print(
            f"   comm_tables_main_df shape: "
            f"{comm_tables_main_df.shape if isinstance(comm_tables_main_df, pd.DataFrame) else None}"
        )

        # -------------------------------------------------------------------
        # Step 1: Read raw first sheet
        # -------------------------------------------------------------------
        print("🟦 Step 1: Read raw first sheet")
        df = read_raw_first_sheet(file_path, header=None)
        print(f"   raw shape: {df.shape}")

        if df is None or df.empty:
            print("🔴 Raw AGBus sheet returned no data.")
            return None

        # -------------------------------------------------------------------
        # Step 2: Find the header row
        # -------------------------------------------------------------------
        print("🟦 Step 2: Find 'Earning Year' row")
        header_row_idx = find_row_index_by_first_column_value(df, "Earning Year")
        print(f"   header_row_idx: {header_row_idx}")

        if header_row_idx is None:
            print("🔴 'Earning Year' row not found.")
            return None

        # -------------------------------------------------------------------
        # Step 3: Set header from target row
        # -------------------------------------------------------------------
        print("🟦 Step 3: Set header from target row")
        df = set_header_from_row(df, header_row_idx)
        print(f"   shape after header set: {df.shape}")

        # -------------------------------------------------------------------
        # Step 4: Normalize headers
        # -------------------------------------------------------------------
        print("🟦 Step 4: Normalize column headers")
        df = normalize_column_headers(df)
        print(f"   columns after normalization: {list(df.columns)}")

        # -------------------------------------------------------------------
        # Step 5: Drop likely blank rows
        # -------------------------------------------------------------------
        print("🟦 Step 5: Drop likely blank rows")
        df = drop_blank_rows_by_column_position(df, 1)
        df = drop_fully_blank_rows(df)
        print(f"   shape after blank-row cleanup: {df.shape}")

        if df.empty:
            print("🔴 No usable rows remain after initial cleanup.")
            return None

        # -------------------------------------------------------------------
        # Step 6: Find source columns
        # -------------------------------------------------------------------
        print("🟦 Step 6: Detect source columns for normalized mapping")

        def f5_1_1_find_matching_column(columns_list, candidate_names):
            """
            Match a real source column by comparing normalized text.
            """
            normalized_map = {
                normalize_text(col): col
                for col in columns_list
            }

            # 1) Exact normalized candidate match
            for candidate in candidate_names:
                candidate_norm = normalize_text(candidate)
                if candidate_norm in normalized_map:
                    return normalized_map[candidate_norm]

            # 2) Relaxed contains match
            for candidate in candidate_names:
                candidate_norm = normalize_text(candidate)
                for norm_col, original_col in normalized_map.items():
                    if candidate_norm and candidate_norm in norm_col:
                        return original_col

            return None

        contract_col_candidates = [
            "Contract Number",
            "Contract No",
            "Contract",
            "Policy Number",
            "Policy No",
            "Policy",
            "Member Number",
            "Member No",
            "Membership Number",
            "Membership No",
            "Account Number",
            "Account No",
        ]

        total_commission_col_candidates = [
            "Total Commission Payable",
            "Commission",
            "Commission Amount",
            "Net Commission",
            "Total Due",

        ]

        planner_col_candidates = [
            "Planner",
            "Broker",
            "Broker Name",
            "Agency",
            "Agency Name",
            "Intermediary",
            "Intermediary Name",
            "Company",
            "Company Name",
            "Practice",
            "Practice Name",
        ]

        client_name_col_candidates = [
            "Client Name",
            "Client",
            "Customer Name",
            "Customer",
            "Policyholder",
            "Policy Holder",
            "Member Name",
            "Member",
            "Insured Name",
            "Name",
        ]

        contract_col = f5_1_1_find_matching_column(list(df.columns), contract_col_candidates)
        total_commission_col = f5_1_1_find_matching_column(list(df.columns), total_commission_col_candidates)
        planner_col = f5_1_1_find_matching_column(list(df.columns), planner_col_candidates)
        client_name_col = f5_1_1_find_matching_column(list(df.columns), client_name_col_candidates)

        print(f"   contract_col        : {contract_col}")
        print(f"   total_commission_col: {total_commission_col}")
        print(f"   planner_col         : {planner_col}")
        print(f"   client_name_col     : {client_name_col}")

        if contract_col is None:
            print("🔴 Could not identify a contract number column in the AGBus file.")
            return None

        if total_commission_col is None:
            print("🔴 Could not identify a total commission column in the AGBus file.")
            return None


        # -------------------------------------------------------------------
        # Step 7: Build normalized DataFrame
        # -------------------------------------------------------------------
        print("🟦 Step 7: Build normalized output DataFrame")

        normalized_df = pd.DataFrame()

        # 7.1 Contract Number
        normalized_df["contract_number"] = (
            df[contract_col]
            .apply(normalize_text_preserve_case)
            .astype(str)
            .str.strip()
        )

        # 7.2 Total Commission
        total_series = (
            df[total_commission_col]
            .astype(str)
            .str.replace(",", "", regex=False)
            .str.replace("R", "", regex=False)
            .str.strip()
        )
        normalized_df["total_commission"] = pd.to_numeric(total_series, errors="coerce")

        # 7.3 Client Name
        if client_name_col is not None:
            normalized_df["client_name"] = (
                df[client_name_col]
                .apply(normalize_text_preserve_case)
                .replace("", "tbc")
            )
        else:
            normalized_df["client_name"] = "tbc"

        # 7.4 Planner
        if planner_col is not None:
            normalized_df["planner"] = (
                df[planner_col]
                .apply(normalize_text_preserve_case)
                .replace("", "UNIFY (PTY) LTD")
            )
        else:
            normalized_df["planner"] = "UNIFY (PTY) LTD"

        # 7.5 Fixed standardized fields
        normalized_df["product_house"] = "agbus"
        normalized_df["commission_month"] = comm_month

        print(f"   normalized_df shape before row filters: {normalized_df.shape}")
        print(f"   normalized_df columns: {list(normalized_df.columns)}")

        # -------------------------------------------------------------------
        # Step 8: Filter invalid rows
        # -------------------------------------------------------------------
        print("🟦 Step 8: Filter invalid normalized rows")

        normalized_df = normalized_df[
            normalized_df["contract_number"].apply(lambda x: normalize_text(x) != "")
        ].copy()

        normalized_df = normalized_df[
            normalized_df["total_commission"].notna()
        ].copy()

        normalized_df = normalized_df.reset_index(drop=True)

        print(f"   normalized_df shape after row filters: {normalized_df.shape}")

        # -------------------------------------------------------------------
        # Step 9: Reorder final columns
        # -------------------------------------------------------------------
        print("🟦 Step 9: Reorder final columns")

        normalized_df = normalized_df[
            [
                "client_name",
                "product_house",
                "commission_month",
                "contract_number",
                "total_commission",
                "planner",
            ]
        ].copy()

        print("✅ AGBus normalized read complete")
        print("   final shape:", normalized_df.shape)
        print("   final preview:")
        print(normalized_df.head(10))

        return normalized_df

    except Exception as e:
        print(f"🔴 Error in read_agbus_df: {e}")
        print(traceback.format_exc())
        return None

# ---------------------------------------------------------------------------
# 5.1.2 Function: Write Official DF To Selected File / processed_data Sheet
# ---------------------------------------------------------------------------
def f5_1_2_write_official_df_to_selected_file(file_path, df):
    try:
        print("------------------------------------------------------------")
        print("🟦 5.1.2 Write Official DF To Selected File")
        print(f"   file_path: {file_path}")

        if df is None or not isinstance(df, pd.DataFrame) or df.empty:
            print("🟠 DF is empty or invalid. Nothing to write.")
            return "no_data"

        file_ext = os.path.splitext(file_path)[1].lower()
        print(f"   file_ext: {file_ext}")

        # ---------------------------------------------------------------
        # Step 1: Excel source files -> existing logic
        # ---------------------------------------------------------------
        if file_ext in [".xlsx", ".xls", ".xlsm"]:
            print("🟦 Step 1: Open selected Excel workbook")
            target_wb = xw.Book(file_path)

        # ---------------------------------------------------------------
        # Step 2: PDF source files -> create companion output workbook
        # ---------------------------------------------------------------
        elif file_ext == ".pdf":
            print("🟦 Step 1: Source is PDF, create companion processed workbook")
            output_file_path = os.path.splitext(file_path)[0] + "_processed.xlsx"
            print(f"   output_file_path: {output_file_path}")

            target_wb = xw.Book()
            target_wb.save(output_file_path)

        else:
            print(f"🔴 Unsupported file extension for official write: {file_ext}")
            return "failed"

        # ---------------------------------------------------------------
        # Step 3: Get or create processed_data sheet
        # ---------------------------------------------------------------
        processed_sheet_name = "processed_data"
        existing_sheet_names = [s.name.lower() for s in target_wb.sheets]

        if processed_sheet_name.lower() in existing_sheet_names:
            sh_processed_data = None
            for s in target_wb.sheets:
                if s.name.strip().lower() == processed_sheet_name.lower():
                    sh_processed_data = s
                    break
        else:
            sh_processed_data = target_wb.sheets.add(processed_sheet_name)

        # ---------------------------------------------------------------
        # Step 4: Clear and write
        # ---------------------------------------------------------------
        sh_processed_data.range("A1:ZZ4000").clear_contents()
        sh_processed_data.range("A1").options(index=False, header=True).value = df

        try:
            sh_processed_data.autofit()
        except Exception as autofit_error:
            print(f"   ⚠️ Autofit skipped: {autofit_error}")

        # ---------------------------------------------------------------
        # Step 5: Save and close
        # ---------------------------------------------------------------
        target_wb.save()
        target_wb.close()

        print("✅ Official DF write complete")
        return "success"

    except Exception as e:
        print(f"🔴 Error in f5_1_2_write_official_df_to_selected_file: {e}")
        print(traceback.format_exc())
        return "failed"

# ---------------------------------------------------------------------------
# 5.1.2 Reader: AllanGray_Unify
# ---------------------------------------------------------------------------
@register_reader("AllanGray_Unify")
def read_allangray_unify_df(file_path, **kwargs):
    """
    Reader for AllanGray_Unify commission statement files.

    Expected output columns:
    - client_name
    - product_house
    - commission_month
    - contract_number
    - total_commission
    - planner
    """
    try:
        print("------------------------------------------------------------")
        print("🧾 5.1.2 Reader: read_allangray_unify_df")
        print(f"   file_path: {file_path}")

        # -------------------------------------------------------------------
        # Step 0: Get passed variables
        # -------------------------------------------------------------------
        print("🟦 Step 0: Get passed variables")
        comm_month = kwargs.get("comm_month")

        print(f"   comm_month: {comm_month}")

        # -------------------------------------------------------------------
        # Step 1: Read raw first sheet
        # -------------------------------------------------------------------
        print("🟦 Step 1: Read raw first sheet")
        df = read_raw_first_sheet(file_path, header=None)
        print(f"   raw shape: {df.shape}")

        if df is None or df.empty:
            print("🔴 Raw AllanGray_Unify sheet returned no data.")
            return None

        # -------------------------------------------------------------------
        # Step 2: Find the header row
        # -------------------------------------------------------------------
        print("🟦 Step 2: Find 'Payment date' row")
        header_row_idx = find_row_index_by_first_column_value(df, "Payment date")
        print(f"   header_row_idx: {header_row_idx}")

        if header_row_idx is None:
            print("🔴 'Payment date' row not found.")
            return None

        # -------------------------------------------------------------------
        # Step 3: Set header from target row
        # -------------------------------------------------------------------
        print("🟦 Step 3: Set header from target row")
        df = set_header_from_row(df, header_row_idx)
        print(f"   shape after header set: {df.shape}")

        # -------------------------------------------------------------------
        # Step 4: Normalize headers
        # -------------------------------------------------------------------
        print("🟦 Step 4: Normalize column headers")
        df = normalize_column_headers(df)
        print(f"   columns after normalization: {list(df.columns)}")

        # -------------------------------------------------------------------
        # Step 5: Drop blank rows
        # -------------------------------------------------------------------
        print("🟦 Step 5: Drop blank rows")
        df = drop_fully_blank_rows(df)
        print(f"   shape after blank-row cleanup: {df.shape}")

        # -------------------------------------------------------------------
        # Step 6: Find source columns
        # -------------------------------------------------------------------
        print("🟦 Step 6: Detect source columns")

        client_name_col = find_matching_column(df.columns, ["Client name"])
        contract_col = find_matching_column(df.columns, ["Account number"])
        total_commission_col = find_matching_column(df.columns, ["Fees earned"])
        planner_col = find_matching_column(df.columns, ["Account Adviser Name"])

        print(f"   client_name_col       : {client_name_col}")
        print(f"   contract_col          : {contract_col}")
        print(f"   total_commission_col  : {total_commission_col}")
        print(f"   planner_col           : {planner_col}")

        if contract_col is None or total_commission_col is None:
            print("🔴 Required AllanGray_Unify columns could not be identified.")
            return None

        # -------------------------------------------------------------------
        # Step 7: Build normalized output df
        # -------------------------------------------------------------------
        print("🟦 Step 7: Build normalized output df")
        normalized_df = pd.DataFrame()

        if client_name_col:
            normalized_df["client_name"] = df[client_name_col].apply(normalize_text_preserve_case)
        else:
            normalized_df["client_name"] = ""

        normalized_df["contract_number"] = df[contract_col].apply(normalize_text_preserve_case)
        normalized_df["total_commission"] = coerce_series_to_numeric(df[total_commission_col])

        if planner_col:
            normalized_df["planner"] = (
                df[planner_col]
                .apply(normalize_text_preserve_case)
                .replace("", "UNIFY (PTY) LTD")
            )
        else:
            normalized_df["planner"] = "UNIFY (PTY) LTD"

        normalized_df["product_house"] = "allangray_unify"
        normalized_df["commission_month"] = comm_month

        # -------------------------------------------------------------------
        # Step 8: Filter invalid rows
        # -------------------------------------------------------------------
        print("🟦 Step 8: Filter invalid rows")

        normalized_df = normalized_df[
            normalized_df["contract_number"].apply(lambda x: normalize_text(x) != "")
        ].copy()

        normalized_df = normalized_df[
            normalized_df["total_commission"].notna()
        ].copy()

        normalized_df = normalized_df.reset_index(drop=True)
        print(f"   normalized_df shape after filters: {normalized_df.shape}")

        if normalized_df.empty:
            print("🔴 No usable AllanGray_Unify rows remain after filtering.")
            return None

        # -------------------------------------------------------------------
        # Step 9: Reorder final columns
        # -------------------------------------------------------------------
        print("🟦 Step 9: Reorder final columns")
        normalized_df = normalized_df[
            [
                "client_name",
                "product_house",
                "commission_month",
                "contract_number",
                "total_commission",
                "planner",
            ]
        ].copy()

        print("✅ AllanGray_Unify normalized read complete")
        print(normalized_df.head(10))

        return normalized_df

    except Exception as e:
        print(f"🔴 Error in read_allangray_unify_df: {e}")
        print(traceback.format_exc())
        return None


# ---------------------------------------------------------------------------
# 5.1.3 Reader: Brolink
# ---------------------------------------------------------------------------
@register_reader("Brolink")
def read_brolink_df(file_path, **kwargs):
    """
    Reader for Brolink commission statement files.
    """
    try:
        print("------------------------------------------------------------")
        print("🧾 5.1.3 Reader: read_brolink_df")
        print(f"   file_path: {file_path}")

        # -------------------------------------------------------------------
        # Step 0: Get passed variables
        # -------------------------------------------------------------------
        print("🟦 Step 0: Get passed variables")
        comm_month = kwargs.get("comm_month")
        print(f"   comm_month: {comm_month}")

        # -------------------------------------------------------------------
        # Step 1: Read raw first sheet
        # -------------------------------------------------------------------
        print("🟦 Step 1: Read raw first sheet")
        df = read_raw_first_sheet(file_path, header=None)
        print(f"   raw shape: {df.shape}")

        if df is None or df.empty:
            print("🔴 Raw Brolink sheet returned no data.")
            return None

        # -------------------------------------------------------------------
        # Step 2: Find header row
        # -------------------------------------------------------------------
        print("🟦 Step 2: Find 'Policy No' row")
        header_row_idx = find_row_index_by_first_column_value(df, "Policy No")
        print(f"   header_row_idx: {header_row_idx}")

        if header_row_idx is None:
            print("🔴 'Policy No' row not found.")
            return None

        # -------------------------------------------------------------------
        # Step 3: Set header from row
        # -------------------------------------------------------------------
        print("🟦 Step 3: Set header from target row")
        df = set_header_from_row(df, header_row_idx)
        print(f"   shape after header set: {df.shape}")

        # -------------------------------------------------------------------
        # Step 4: Normalize headers
        # -------------------------------------------------------------------
        print("🟦 Step 4: Normalize column headers")
        df = normalize_column_headers(df)
        print(f"   columns after normalization: {list(df.columns)}")

        # -------------------------------------------------------------------
        # Step 5: Drop blank rows
        # -------------------------------------------------------------------
        print("🟦 Step 5: Drop blank rows")
        df = drop_fully_blank_rows(df)
        print(f"   shape after blank-row cleanup: {df.shape}")

        # -------------------------------------------------------------------
        # Step 6: Find source columns
        # -------------------------------------------------------------------
        print("🟦 Step 6: Detect source columns")

        client_name_col = find_matching_column(df.columns, ["Client Name"])
        contract_col = find_matching_column(df.columns, ["Policy No"])
        total_commission_col = find_matching_column(df.columns, ["Total Payable Ex Vat", "Total Payable"])
        planner_col = find_matching_column(df.columns, ["Broker Contact Name", "Agent Name", "Marketer Name"])

        print(f"   client_name_col       : {client_name_col}")
        print(f"   contract_col          : {contract_col}")
        print(f"   total_commission_col  : {total_commission_col}")
        print(f"   planner_col           : {planner_col}")

        if contract_col is None or total_commission_col is None:
            print("🔴 Required Brolink columns could not be identified.")
            return None

        # -------------------------------------------------------------------
        # Step 7: Build normalized output df
        # -------------------------------------------------------------------
        print("🟦 Step 7: Build normalized output df")
        normalized_df = pd.DataFrame()

        if client_name_col:
            normalized_df["client_name"] = df[client_name_col].apply(normalize_text_preserve_case)
        else:
            normalized_df["client_name"] = ""

        normalized_df["contract_number"] = df[contract_col].apply(normalize_text_preserve_case)
        normalized_df["total_commission"] = coerce_series_to_numeric(df[total_commission_col])

        if planner_col:
            normalized_df["planner"] = (
                df[planner_col]
                .apply(normalize_text_preserve_case)
                .replace("", "UNIFY (PTY) LTD")
            )
        else:
            normalized_df["planner"] = "UNIFY (PTY) LTD"

        normalized_df["product_house"] = "brolink"
        normalized_df["commission_month"] = comm_month

        # -------------------------------------------------------------------
        # Step 8: Filter invalid rows
        # -------------------------------------------------------------------
        print("🟦 Step 8: Filter invalid rows")
        normalized_df = normalized_df[
            normalized_df["contract_number"].apply(lambda x: normalize_text(x) != "")
        ].copy()

        normalized_df = normalized_df[
            normalized_df["total_commission"].notna()
        ].copy()

        normalized_df = normalized_df.reset_index(drop=True)
        print(f"   normalized_df shape after filters: {normalized_df.shape}")

        if normalized_df.empty:
            print("🔴 No usable Brolink rows remain after filtering.")
            return None

        # -------------------------------------------------------------------
        # Step 9: Reorder final columns
        # -------------------------------------------------------------------
        print("🟦 Step 9: Reorder final columns")
        normalized_df = normalized_df[
            [
                "client_name",
                "product_house",
                "commission_month",
                "contract_number",
                "total_commission",
                "planner",
            ]
        ].copy()

        print("✅ Brolink normalized read complete")
        print(normalized_df.head(10))

        return normalized_df

    except Exception as e:
        print(f"🔴 Error in read_brolink_df: {e}")
        print(traceback.format_exc())
        return None


# ---------------------------------------------------------------------------
# 5.1.4 Reader: Disclife
# ---------------------------------------------------------------------------
@register_reader("Disclife")
def read_disclife_df(file_path, **kwargs):
    """
    Reader for Disclife commission statement files.
    """
    try:
        print("------------------------------------------------------------")
        print("🧾 5.1.4 Reader: read_disclife_df")
        print(f"   file_path: {file_path}")

        # -------------------------------------------------------------------
        # Step 0: Get passed variables
        # -------------------------------------------------------------------
        print("🟦 Step 0: Get passed variables")
        comm_month = kwargs.get("comm_month")
        print(f"   comm_month: {comm_month}")

        # -------------------------------------------------------------------
        # Step 1: Read raw first sheet
        # -------------------------------------------------------------------
        print("🟦 Step 1: Read raw first sheet")
        df = read_raw_first_sheet(file_path, header=None)
        print(f"   raw shape: {df.shape}")

        if df is None or df.empty:
            print("🔴 Raw Disclife sheet returned no data.")
            return None

        # -------------------------------------------------------------------
        # Step 2: Find header row
        # -------------------------------------------------------------------
        print("🟦 Step 2: Find 'Broker No' row")
        header_row_idx = find_row_index_by_first_column_value(df, "Broker No")
        print(f"   header_row_idx: {header_row_idx}")

        if header_row_idx is None:
            print("🔴 'Broker No' row not found.")
            return None

        # -------------------------------------------------------------------
        # Step 3: Set header from row
        # -------------------------------------------------------------------
        print("🟦 Step 3: Set header from target row")
        df = set_header_from_row(df, header_row_idx)
        print(f"   shape after header set: {df.shape}")

        # -------------------------------------------------------------------
        # Step 4: Normalize headers
        # -------------------------------------------------------------------
        print("🟦 Step 4: Normalize column headers")
        df = normalize_column_headers(df)
        print(f"   columns after normalization: {list(df.columns)}")

        # -------------------------------------------------------------------
        # Step 5: Drop blank rows
        # -------------------------------------------------------------------
        print("🟦 Step 5: Drop blank rows")
        df = drop_fully_blank_rows(df)
        print(f"   shape after blank-row cleanup: {df.shape}")

        # -------------------------------------------------------------------
        # Step 6: Find source columns
        # -------------------------------------------------------------------
        print("🟦 Step 6: Detect source columns")

        client_name_col = find_matching_column(df.columns, ["Policy Name"])
        contract_col = find_matching_column(df.columns, ["Policy No"])
        total_commission_col = find_matching_column(df.columns, ["Total Amount", "Final amount(exclude vat)"])
        planner_col = find_matching_column(df.columns, ["Broker Name"])

        print(f"   client_name_col       : {client_name_col}")
        print(f"   contract_col          : {contract_col}")
        print(f"   total_commission_col  : {total_commission_col}")
        print(f"   planner_col           : {planner_col}")

        if contract_col is None or total_commission_col is None:
            print("🔴 Required Disclife columns could not be identified.")
            return None

        # -------------------------------------------------------------------
        # Step 7: Build normalized output df
        # -------------------------------------------------------------------
        print("🟦 Step 7: Build normalized output df")
        normalized_df = pd.DataFrame()

        if client_name_col:
            normalized_df["client_name"] = df[client_name_col].apply(normalize_text_preserve_case)
        else:
            normalized_df["client_name"] = ""

        normalized_df["contract_number"] = df[contract_col].apply(normalize_text_preserve_case)
        normalized_df["total_commission"] = coerce_series_to_numeric(df[total_commission_col])

        if planner_col:
            normalized_df["planner"] = (
                df[planner_col]
                .apply(normalize_text_preserve_case)
                .replace("", "UNIFY (PTY) LTD")
            )
        else:
            normalized_df["planner"] = "UNIFY (PTY) LTD"

        normalized_df["product_house"] = "disclife"
        normalized_df["commission_month"] = comm_month

        # -------------------------------------------------------------------
        # Step 8: Filter invalid rows
        # -------------------------------------------------------------------
        print("🟦 Step 8: Filter invalid rows")
        normalized_df = normalized_df[
            normalized_df["contract_number"].apply(lambda x: normalize_text(x) != "")
        ].copy()

        normalized_df = normalized_df[
            normalized_df["total_commission"].notna()
        ].copy()

        normalized_df = normalized_df.reset_index(drop=True)
        print(f"   normalized_df shape after filters: {normalized_df.shape}")

        if normalized_df.empty:
            print("🔴 No usable Disclife rows remain after filtering.")
            return None

        # -------------------------------------------------------------------
        # Step 9: Reorder final columns
        # -------------------------------------------------------------------
        print("🟦 Step 9: Reorder final columns")
        normalized_df = normalized_df[
            [
                "client_name",
                "product_house",
                "commission_month",
                "contract_number",
                "total_commission",
                "planner",
            ]
        ].copy()

        print("✅ Disclife normalized read complete")
        print(normalized_df.head(10))

        return normalized_df

    except Exception as e:
        print(f"🔴 Error in read_disclife_df: {e}")
        print(traceback.format_exc())
        return None

# ---------------------------------------------------------------------------
# 5.1.5 Reader: Dischealth
# ---------------------------------------------------------------------------
@register_reader("Dischealth")
def read_dischealth_df(file_path, **kwargs):
    """
    Reader for Dischealth commission statement files.

    Note:
    The sample file has two blank columns after 'Comm Month'.
    The first of those carries the actual commission amount in the sample.
    """
    try:
        print("------------------------------------------------------------")
        print("🧾 5.1.5 Reader: read_dischealth_df")
        print(f"   file_path: {file_path}")

        # -------------------------------------------------------------------
        # Step 0: Get passed variables
        # -------------------------------------------------------------------
        print("🟦 Step 0: Get passed variables")
        comm_month = kwargs.get("comm_month")
        print(f"   comm_month: {comm_month}")

        # -------------------------------------------------------------------
        # Step 1: Read raw first sheet
        # -------------------------------------------------------------------
        print("🟦 Step 1: Read raw first sheet")
        df = read_raw_first_sheet(file_path, header=None)
        print(f"   raw shape: {df.shape}")

        if df is None or df.empty:
            print("🔴 Raw Dischealth sheet returned no data.")
            return None

        # -------------------------------------------------------------------
        # Step 2: Find header row
        # -------------------------------------------------------------------
        print("🟦 Step 2: Find 'Internal Company' row")
        header_row_idx = find_row_index_by_first_column_value(df, "Internal Company")
        print(f"   header_row_idx: {header_row_idx}")

        if header_row_idx is None:
            print("🔴 'Internal Company' row not found.")
            return None

        # -------------------------------------------------------------------
        # Step 3: Set header from row
        # -------------------------------------------------------------------
        print("🟦 Step 3: Set header from target row")
        df = set_header_from_row(df, header_row_idx)
        print(f"   shape after header set: {df.shape}")

        # -------------------------------------------------------------------
        # Step 4: Normalize headers
        # -------------------------------------------------------------------
        print("🟦 Step 4: Normalize column headers")
        df = normalize_column_headers(df)
        print(f"   columns after normalization: {list(df.columns)}")

        # -------------------------------------------------------------------
        # Step 5: Drop blank rows
        # -------------------------------------------------------------------
        print("🟦 Step 5: Drop blank rows")
        df = drop_fully_blank_rows(df)
        print(f"   shape after blank-row cleanup: {df.shape}")

        # -------------------------------------------------------------------
        # Step 6: Find source columns
        # -------------------------------------------------------------------
        print("🟦 Step 6: Detect source columns")

        client_name_col = find_matching_column(df.columns, ["Member Name"])
        contract_col = find_matching_column(df.columns, ["Member No."])
        planner_col = find_matching_column(df.columns, ["Broker Name"])
        comm_month_source_col = find_matching_column(df.columns, ["Comm Month"])

        total_commission_col = None

        # Prefer the first column immediately after Comm Month
        if comm_month_source_col in df.columns:
            comm_month_idx = list(df.columns).index(comm_month_source_col)

            if comm_month_idx + 1 < len(df.columns):
                first_after_comm_month = df.columns[comm_month_idx + 1]

                temp_numeric = coerce_series_to_numeric(df[first_after_comm_month])
                if temp_numeric.notna().sum() > 0:
                    total_commission_col = first_after_comm_month

            # if first one fails, try second column after Comm Month
            if total_commission_col is None and comm_month_idx + 2 < len(df.columns):
                second_after_comm_month = df.columns[comm_month_idx + 2]

                temp_numeric = coerce_series_to_numeric(df[second_after_comm_month])
                if temp_numeric.notna().sum() > 0:
                    total_commission_col = second_after_comm_month

        # fallback to named columns if needed
        if total_commission_col is None:
            total_commission_col = find_matching_column(
                df.columns,
                ["Amount Due", "Total Amount", "Commission Amount"]
            )

        print(f"   client_name_col       : {client_name_col}")
        print(f"   contract_col          : {contract_col}")
        print(f"   total_commission_col  : {total_commission_col}")
        print(f"   planner_col           : {planner_col}")

        if contract_col is None or total_commission_col is None:
            print("🔴 Required Dischealth columns could not be identified.")
            return None

        # -------------------------------------------------------------------
        # Step 7: Build normalized output df
        # -------------------------------------------------------------------
        print("🟦 Step 7: Build normalized output df")
        normalized_df = pd.DataFrame()

        if client_name_col:
            normalized_df["client_name"] = df[client_name_col].apply(normalize_text_preserve_case)
        else:
            normalized_df["client_name"] = ""

        normalized_df["contract_number"] = df[contract_col].apply(normalize_text_preserve_case)
        normalized_df["total_commission"] = coerce_series_to_numeric(df[total_commission_col])

        if planner_col:
            normalized_df["planner"] = (
                df[planner_col]
                .apply(normalize_text_preserve_case)
                .replace("", "UNIFY (PTY) LTD")
            )
        else:
            normalized_df["planner"] = "UNIFY (PTY) LTD"

        normalized_df["product_house"] = "dischealth"
        normalized_df["commission_month"] = comm_month

        # -------------------------------------------------------------------
        # Step 8: Filter invalid rows
        # -------------------------------------------------------------------
        print("🟦 Step 8: Filter invalid rows")
        normalized_df = normalized_df[
            normalized_df["contract_number"].apply(lambda x: normalize_text(x) != "")
        ].copy()

        normalized_df = normalized_df[
            normalized_df["total_commission"].notna()
        ].copy()

        normalized_df = normalized_df.reset_index(drop=True)
        print(f"   normalized_df shape after filters: {normalized_df.shape}")

        if normalized_df.empty:
            print("🔴 No usable Dischealth rows remain after filtering.")
            return None

        # -------------------------------------------------------------------
        # Step 9: Reorder final columns
        # -------------------------------------------------------------------
        print("🟦 Step 9: Reorder final columns")
        normalized_df = normalized_df[
            [
                "client_name",
                "product_house",
                "commission_month",
                "contract_number",
                "total_commission",
                "planner",
            ]
        ].copy()

        print("✅ Dischealth normalized read complete")
        print(normalized_df.head(10))

        return normalized_df

    except Exception as e:
        print(f"🔴 Error in read_dischealth_df: {e}")
        print(traceback.format_exc())
        return None

# ---------------------------------------------------------------------------
# 5.1.6 Reader: Discgap
# ---------------------------------------------------------------------------
@register_reader("Discgap")
def read_discgap_df(file_path, **kwargs):
    """
    Reader for Discgap commission statement files.

    Goal:
    - read the raw Discgap sheet correctly
    - identify the true header row
    - clean the raw data
    - map the source columns into a normalized structure

    Returns normalized columns:
    - client_name
    - product_house
    - commission_month
    - contract_number
    - total_commission
    - planner
    """
    try:
        print("------------------------------------------------------------")
        print("🧾 5.1.3 Reader: read_discgap_df")
        print(f"   file_path: {file_path}")

        # -------------------------------------------------------------------
        # Step 0: Get Variables
        # -------------------------------------------------------------------
        print("🟦 Step 0: Get passed variables")
        comm_month = kwargs.get("comm_month")
        comm_tables_main_df = kwargs.get("comm_tables_main_df")

        print(f"   comm_month: {comm_month}")
        print(
            f"   comm_tables_main_df shape: "
            f"{comm_tables_main_df.shape if isinstance(comm_tables_main_df, pd.DataFrame) else None}"
        )

        # -------------------------------------------------------------------
        # Step 1: Read raw first sheet
        # -------------------------------------------------------------------
        print("🟦 Step 1: Read raw first sheet")
        df = read_raw_first_sheet(file_path, header=None)
        print(f"   raw shape: {df.shape}")

        if df is None or df.empty:
            print("🔴 Raw Discgap sheet returned no data.")
            return None

        # -------------------------------------------------------------------
        # Step 2: Find the header row
        # -------------------------------------------------------------------
        print("🟦 Step 2: Find 'Internal Company' row")
        header_row_idx = find_row_index_by_first_column_value(df, "Internal Company")
        print(f"   header_row_idx: {header_row_idx}")

        if header_row_idx is None:
            print("🔴 'Internal Company' row not found.")
            return None

        # -------------------------------------------------------------------
        # Step 3: Set header from target row
        # -------------------------------------------------------------------
        print("🟦 Step 3: Set header from target row")
        df = set_header_from_row(df, header_row_idx)
        print(f"   shape after header set: {df.shape}")

        # -------------------------------------------------------------------
        # Step 4: Normalize headers
        # -------------------------------------------------------------------
        print("🟦 Step 4: Normalize column headers")
        df = normalize_column_headers(df)
        print(f"   columns after normalization: {list(df.columns)}")

        # -------------------------------------------------------------------
        # Step 5: Drop likely blank rows
        # -------------------------------------------------------------------
        print("🟦 Step 5: Drop likely blank rows")
        df = drop_blank_rows_by_column_position(df, 1)
        df = drop_fully_blank_rows(df)
        print(f"   shape after blank-row cleanup: {df.shape}")

        if df.empty:
            print("🔴 No usable rows remain after initial cleanup.")
            return None

        # -------------------------------------------------------------------
        # Step 6: Find source columns
        # -------------------------------------------------------------------
        print("🟦 Step 6: Detect source columns for normalized mapping")

        def f5_1_3_find_matching_column(columns_list, candidate_names):
            """
            Match a real source column by comparing normalized text.
            """
            normalized_map = {
                normalize_text(col): col
                for col in columns_list
            }

            # 1) Exact normalized candidate match
            for candidate in candidate_names:
                candidate_norm = normalize_text(candidate)
                if candidate_norm in normalized_map:
                    return normalized_map[candidate_norm]

            # 2) Relaxed contains match
            for candidate in candidate_names:
                candidate_norm = normalize_text(candidate)
                for norm_col, original_col in normalized_map.items():
                    if candidate_norm and candidate_norm in norm_col:
                        return original_col

            return None

        contract_col_candidates = [
            "Member No.",
            "Member No",
            "Member Number",
            "Membership Number",
            "Membership No",
            "Policy Number",
            "Policy No",
            "Contract Number",
            "Contract No",
            "Account Number",
            "Account No",
        ]

        client_name_col_candidates = [
            "Member Name",
            "Client Name",
            "Client",
            "Customer Name",
            "Policyholder",
            "Policy Holder",
            "Name",
        ]

        planner_col_candidates = [
            "Broker Name",
            "Planner",
            "Broker",
            "Agent",
            "Intermediary",
            "Intermediary Name",
        ]

        comm_month_col_candidates = [
            "Comm Month",
            "Commission Month",
        ]

        contract_col = f5_1_3_find_matching_column(list(df.columns), contract_col_candidates)
        client_name_col = f5_1_3_find_matching_column(list(df.columns), client_name_col_candidates)
        planner_col = f5_1_3_find_matching_column(list(df.columns), planner_col_candidates)
        detected_comm_month_col = f5_1_3_find_matching_column(list(df.columns), comm_month_col_candidates)

        print(f"   contract_col           : {contract_col}")
        print(f"   client_name_col        : {client_name_col}")
        print(f"   planner_col            : {planner_col}")
        print(f"   detected_comm_month_col: {detected_comm_month_col}")

        if contract_col is None:
            print("🔴 Could not identify a contract number column in the Discgap file.")
            return None

        # -------------------------------------------------------------------
        # Step 6.1: Identify total commission column
        # -------------------------------------------------------------------
        print("🟦 Step 6.1: Detect total commission column")

        total_commission_col = None
        total_commission_col_idx = None

        # In the uploaded Discgap style, the true commission amount sits in the
        # first numeric column immediately after 'Comm Month', even though there
        # are blank header labels before the later named columns.
        if detected_comm_month_col is not None:
            comm_month_idx = list(df.columns).index(detected_comm_month_col)

            for idx in range(comm_month_idx + 1, min(comm_month_idx + 4, len(df.columns))):
                numeric_test = pd.to_numeric(df.iloc[:, idx], errors="coerce")

                if numeric_test.notna().sum() >= max(1, int(len(df) * 0.5)):
                    total_commission_col = df.columns[idx]
                    total_commission_col_idx = idx
                    break

        # Fallback if the nearby numeric scan fails
        if total_commission_col_idx is None:
            fallback_total_candidates = [
                "Commission",
                "Commission Amount",
                "Total Commission",
                "Comm VAT",
                "Amount Due",
            ]

            total_commission_col = f5_1_3_find_matching_column(
                list(df.columns),
                fallback_total_candidates
            )

            if total_commission_col is not None:
                total_commission_col_idx = list(df.columns).index(total_commission_col)

        print(f"   total_commission_col    : {total_commission_col}")
        print(f"   total_commission_col_idx: {total_commission_col_idx}")

        if total_commission_col_idx is None:
            print("🔴 Could not identify a total commission column in the Discgap file.")
            return None

        # -------------------------------------------------------------------
        # Step 7: Build normalized DataFrame
        # -------------------------------------------------------------------
        print("🟦 Step 7: Build normalized output DataFrame")

        normalized_df = pd.DataFrame()

        # 7.1 Contract Number
        normalized_df["contract_number"] = (
            df[contract_col]
            .apply(normalize_text_preserve_case)
            .astype(str)
            .str.strip()
        )

        # 7.2 Total Commission
        total_series = (
            df.iloc[:, total_commission_col_idx]
            .astype(str)
            .str.replace(",", "", regex=False)
            .str.replace("R", "", regex=False)
            .str.strip()
        )
        normalized_df["total_commission"] = pd.to_numeric(total_series, errors="coerce")

        # 7.3 Client Name
        if client_name_col is not None:
            normalized_df["client_name"] = (
                df[client_name_col]
                .apply(normalize_text_preserve_case)
                .replace("", "tbc")
            )
        else:
            normalized_df["client_name"] = "tbc"

        # 7.4 Planner
        if planner_col is not None:
            normalized_df["planner"] = (
                df[planner_col]
                .apply(normalize_text_preserve_case)
                .replace("", "tbc")
            )
        else:
            normalized_df["planner"] = "tbc"

        # 7.5 Fixed standardized fields
        normalized_df["product_house"] = "discgap"
        normalized_df["commission_month"] = comm_month

        print(f"   normalized_df shape before row filters: {normalized_df.shape}")
        print(f"   normalized_df columns: {list(normalized_df.columns)}")

        # -------------------------------------------------------------------
        # Step 8: Filter invalid rows
        # -------------------------------------------------------------------
        print("🟦 Step 8: Filter invalid normalized rows")

        normalized_df = normalized_df[
            normalized_df["contract_number"].apply(lambda x: normalize_text(x) != "")
        ].copy()

        normalized_df = normalized_df[
            normalized_df["total_commission"].notna()
        ].copy()

        normalized_df = normalized_df.reset_index(drop=True)

        print(f"   normalized_df shape after row filters: {normalized_df.shape}")

        # -------------------------------------------------------------------
        # Step 9: Reorder final columns
        # -------------------------------------------------------------------
        print("🟦 Step 9: Reorder final columns")

        normalized_df = normalized_df[
            [
                "client_name",
                "product_house",
                "commission_month",
                "contract_number",
                "total_commission",
                "planner",
            ]
        ].copy()

        print("✅ Discgap normalized read complete")
        print("   final shape:", normalized_df.shape)
        print("   final preview:")
        print(normalized_df.head(10))

        return normalized_df

    except Exception as e:
        print(f"🔴 Error in read_discgap_df: {e}")
        print(traceback.format_exc())
        return None

# ---------------------------------------------------------------------------
# 5.1.x Reader: Discinsure
# ---------------------------------------------------------------------------
@register_reader("Discinsure")
def read_discinsure_df(file_path, **kwargs):
    """
    Reader for Discinsure commission statement files.


    Expected output columns:
    - client_name
    - product_house
    - commission_month
    - contract_number
    - total_commission
    - planner
    """
    try:
        print("------------------------------------------------------------")
        print("🧾 5.1.x Reader: read_discinsure_df")
        print(f"   file_path: {file_path}")


        # -------------------------------------------------------------------
        # Step 0: Get passed variables
        # -------------------------------------------------------------------
        print("🟦 Step 0: Get passed variables")
        comm_month = kwargs.get("comm_month")
        print(f"   comm_month: {comm_month}")


        # -------------------------------------------------------------------
        # Step 1: Read raw first sheet
        # -------------------------------------------------------------------
        print("🟦 Step 1: Read raw first sheet")
        df = read_raw_first_sheet(file_path, header=None)
        print(f"   raw shape: {df.shape}")


        if df is None or df.empty:
            print("🔴 Raw Discinsure sheet returned no data.")
            return None


        # -------------------------------------------------------------------
        # Step 2: Find header row
        # -------------------------------------------------------------------
        print("🟦 Step 2: Find 'Broker name' row")
        header_row_idx = find_row_index_by_first_column_value(df, "Broker name")
        print(f"   header_row_idx: {header_row_idx}")


        if header_row_idx is None:
            print("🔴 'Broker name' row not found.")
            return None


        # -------------------------------------------------------------------
        # Step 3: Set header from target row
        # -------------------------------------------------------------------
        print("🟦 Step 3: Set header from target row")
        df = set_header_from_row(df, header_row_idx)
        print(f"   shape after header set: {df.shape}")


        # -------------------------------------------------------------------
        # Step 4: Normalize headers
        # -------------------------------------------------------------------
        print("🟦 Step 4: Normalize column headers")
        df = normalize_column_headers(df)
        print(f"   columns after normalization: {list(df.columns)}")


        # -------------------------------------------------------------------
        # Step 5: Drop blank rows
        # -------------------------------------------------------------------
        print("🟦 Step 5: Drop blank rows")
        df = drop_fully_blank_rows(df)
        print(f"   shape after blank-row cleanup: {df.shape}")


        # -------------------------------------------------------------------
        # Step 6: Detect source columns
        # -------------------------------------------------------------------
        print("🟦 Step 6: Detect source columns")
        client_name_col = find_matching_column(
            df.columns,
            ["Policy Holder", "Policy holder", "Client Name", "Client", "Member Name"]
        )
        contract_col = find_matching_column(
            df.columns,
            ["Policy Number", "Policy No", "Contract Number"]
        )
        total_commission_col = find_matching_column(
            df.columns,
            ["Total Amount", "Comm Amount", "Commission Amount", "Commission"]
        )
        planner_col = find_matching_column(
            df.columns,
            ["Broker name", "Planner", "Broker"]
        )


        print(f"   client_name_col       : {client_name_col}")
        print(f"   contract_col          : {contract_col}")
        print(f"   total_commission_col  : {total_commission_col}")
        print(f"   planner_col           : {planner_col}")


        if contract_col is None or total_commission_col is None:
            print("🔴 Required Discinsure columns could not be identified.")
            return None


        # -------------------------------------------------------------------
        # Step 7: Build normalized output df
        # -------------------------------------------------------------------
        print("🟦 Step 7: Build normalized output df")
        normalized_df = pd.DataFrame()


        if client_name_col:
            normalized_df["client_name"] = df[client_name_col].apply(normalize_text_preserve_case)
        else:
            normalized_df["client_name"] = ""


        normalized_df["product_house"] = "discinsure"
        normalized_df["commission_month"] = comm_month
        normalized_df["contract_number"] = df[contract_col].apply(normalize_text_preserve_case)
        normalized_df["total_commission"] = coerce_series_to_numeric(df[total_commission_col])


        if planner_col:
            normalized_df["planner"] = (
                df[planner_col]
                .apply(normalize_text_preserve_case)
                .replace("", "UNIFY (PTY) LTD")
            )
        else:
            normalized_df["planner"] = "UNIFY (PTY) LTD"


        # -------------------------------------------------------------------
        # Step 8: Filter invalid rows
        # -------------------------------------------------------------------
        print("🟦 Step 8: Filter invalid rows")
        normalized_df = normalized_df[
            normalized_df["contract_number"].apply(lambda x: normalize_text(x) != "")
        ].copy()


        normalized_df = normalized_df[
            normalized_df["total_commission"].notna()
        ].copy()


        normalized_df = normalized_df.reset_index(drop=True)
        print(f"   normalized_df shape after filters: {normalized_df.shape}")


        if normalized_df.empty:
            print("🔴 No usable Discinsure rows remain after filtering.")
            return None


        # -------------------------------------------------------------------
        # Step 9: Reorder final columns
        # -------------------------------------------------------------------
        print("🟦 Step 9: Reorder final columns")
        normalized_df = normalized_df[
            [
                "client_name",
                "product_house",
                "commission_month",
                "contract_number",
                "total_commission",
                "planner",
            ]
        ].copy()


        print("✅ Discinsure normalized read complete")
        print(normalized_df.head(10))


        return normalized_df


    except Exception as e:
        print(f"🔴 Error in read_discinsure_df: {e}")
        print(traceback.format_exc())
        return None

# ---------------------------------------------------------------------------
# 5.1.6 Reader: Ambledown
# ---------------------------------------------------------------------------
@register_reader("Ambledown")
def read_ambledown_df(file_path, **kwargs):
    """
    Reader for Ambledown commission statement files.
    """
    try:
        print("------------------------------------------------------------")
        print("🧾 5.1.6 Reader: read_ambledown_df")
        print(f"   file_path: {file_path}")

        # -------------------------------------------------------------------
        # Step 0: Get passed variables
        # -------------------------------------------------------------------
        print("🟦 Step 0: Get passed variables")
        comm_month = kwargs.get("comm_month")
        print(f"   comm_month: {comm_month}")

        # -------------------------------------------------------------------
        # Step 1: Read raw first sheet
        # -------------------------------------------------------------------
        print("🟦 Step 1: Read raw first sheet")
        df = read_raw_first_sheet(file_path, header=None)
        print(f"   raw shape: {df.shape}")

        if df is None or df.empty:
            print("🔴 Raw Ambledown sheet returned no data.")
            return None

        # -------------------------------------------------------------------
        # Step 2: Find header row
        # -------------------------------------------------------------------
        print("🟦 Step 2: Find 'Policy Number' row")
        header_row_idx = find_row_index_by_first_column_value(df, "Policy Number")
        print(f"   header_row_idx: {header_row_idx}")

        if header_row_idx is None:
            print("🔴 'Policy Number' row not found.")
            return None

        # -------------------------------------------------------------------
        # Step 3: Set header from row
        # -------------------------------------------------------------------
        print("🟦 Step 3: Set header from target row")
        df = set_header_from_row(df, header_row_idx)
        print(f"   shape after header set: {df.shape}")

        # -------------------------------------------------------------------
        # Step 4: Normalize headers
        # -------------------------------------------------------------------
        print("🟦 Step 4: Normalize column headers")
        df = normalize_column_headers(df)
        print(f"   columns after normalization: {list(df.columns)}")

        # -------------------------------------------------------------------
        # Step 5: Drop blank rows
        # -------------------------------------------------------------------
        print("🟦 Step 5: Drop blank rows")
        df = drop_fully_blank_rows(df)
        print(f"   shape after blank-row cleanup: {df.shape}")

        # -------------------------------------------------------------------
        # Step 6: Find source columns
        # -------------------------------------------------------------------
        print("🟦 Step 6: Detect source columns")

        contract_col = find_matching_column(df.columns, ["Policy Number"])
        surname_col = find_matching_column(df.columns, ["Surname"])
        initials_col = find_matching_column(df.columns, ["Initials"])
        total_commission_col = find_matching_column(df.columns, ["Fee Value"])
        planner_col = find_matching_column(df.columns, ["Advisor Name"])

        print(f"   contract_col          : {contract_col}")
        print(f"   surname_col           : {surname_col}")
        print(f"   initials_col          : {initials_col}")
        print(f"   total_commission_col  : {total_commission_col}")
        print(f"   planner_col           : {planner_col}")

        if contract_col is None or total_commission_col is None:
            print("🔴 Required Ambledown columns could not be identified.")
            return None

        # -------------------------------------------------------------------
        # Step 7: Build normalized output df
        # -------------------------------------------------------------------
        print("🟦 Step 7: Build normalized output df")
        normalized_df = pd.DataFrame()

        surname_series = df[surname_col].apply(normalize_text_preserve_case) if surname_col else ""
        initials_series = df[initials_col].apply(normalize_text_preserve_case) if initials_col else ""

        if surname_col and initials_col:
            normalized_df["client_name"] = (
                surname_series.fillna("").astype(str).str.strip() + " " +
                initials_series.fillna("").astype(str).str.strip()
            ).str.strip()
        elif surname_col:
            normalized_df["client_name"] = surname_series
        else:
            normalized_df["client_name"] = ""

        normalized_df["contract_number"] = df[contract_col].apply(normalize_text_preserve_case)
        normalized_df["total_commission"] = coerce_series_to_numeric(df[total_commission_col])

        if planner_col:
            normalized_df["planner"] = (
                df[planner_col]
                .apply(normalize_text_preserve_case)
                .replace("", "UNIFY (PTY) LTD")
            )
        else:
            normalized_df["planner"] = "UNIFY (PTY) LTD"

        normalized_df["product_house"] = "ambledown"
        normalized_df["commission_month"] = comm_month

        # -------------------------------------------------------------------
        # Step 8: Filter invalid rows
        # -------------------------------------------------------------------
        print("🟦 Step 8: Filter invalid rows")
        normalized_df = normalized_df[
            normalized_df["contract_number"].apply(lambda x: normalize_text(x) != "")
        ].copy()

        normalized_df = normalized_df[
            normalized_df["total_commission"].notna()
        ].copy()

        normalized_df = normalized_df.reset_index(drop=True)
        print(f"   normalized_df shape after filters: {normalized_df.shape}")

        if normalized_df.empty:
            print("🔴 No usable Ambledown rows remain after filtering.")
            return None

        # -------------------------------------------------------------------
        # Step 9: Reorder final columns
        # -------------------------------------------------------------------
        print("🟦 Step 9: Reorder final columns")
        normalized_df = normalized_df[
            [
                "client_name",
                "product_house",
                "commission_month",
                "contract_number",
                "total_commission",
                "planner",
            ]
        ].copy()

        print("✅ Ambledown normalized read complete")
        print(normalized_df.head(10))

        return normalized_df

    except Exception as e:
        print(f"🔴 Error in read_ambledown_df: {e}")
        print(traceback.format_exc())
        return None

# ---------------------------------------------------------------------------
# 5.2 Fixed-Header Reader Functions
# ---------------------------------------------------------------------------
@register_reader("Bidvest")
def read_bidvest_df(file_path, **kwargs):
    """
    Reader for Bidvest commission statement files.

    Goal:
    - read the raw Bidvest sheet correctly
    - normalize the source headers
    - clean the raw data
    - detect the main source columns robustly
    - map the source columns into the normalized structure

    Returns normalized columns:
    - client_name
    - product_house
    - commission_month
    - contract_number
    - total_commission
    - planner
    """
    try:
        print("------------------------------------------------------------")
        print("🧾 5.2.1 Reader: read_bidvest_df")
        print(f"   file_path: {file_path}")

        # -------------------------------------------------------------------
        # Step 0: Get passed variables
        # -------------------------------------------------------------------
        print("🟦 Step 0: Get passed variables")
        comm_month = kwargs.get("comm_month")
        comm_tables_main_df = kwargs.get("comm_tables_main_df")

        print(f"   comm_month: {comm_month}")
        print(
            f"   comm_tables_main_df shape: "
            f"{comm_tables_main_df.shape if isinstance(comm_tables_main_df, pd.DataFrame) else None}"
        )

        # -------------------------------------------------------------------
        # Step 1: Read raw first sheet
        # -------------------------------------------------------------------
        print("🟦 Step 1: Read raw first sheet with header row")
        df = read_raw_first_sheet(file_path, header=0)
        print(f"   raw shape: {df.shape}")

        if df is None or df.empty:
            print("🔴 Raw Bidvest sheet returned no data.")
            return None

        print(f"   raw columns before normalization: {list(df.columns)}")

        # -------------------------------------------------------------------
        # Step 2: Normalize column headers
        # -------------------------------------------------------------------
        print("🟦 Step 2: Normalize column headers")
        df = normalize_column_headers(df)
        print(f"   normalized columns: {list(df.columns)}")

        # -------------------------------------------------------------------
        # Step 3: Drop likely blank rows
        # -------------------------------------------------------------------
        print("🟦 Step 3: Drop likely blank rows")
        df = drop_fully_blank_rows(df)
        print(f"   shape after blank-row cleanup: {df.shape}")

        if df.empty:
            print("🔴 No usable rows remain after initial cleanup.")
            return None

        # -------------------------------------------------------------------
        # Step 4: Find source columns
        # -------------------------------------------------------------------
        print("🟦 Step 4: Detect source columns for normalized mapping")

        def f5_2_1_find_matching_column(columns_list, candidate_names):
            """
            Match a real source column by comparing normalized text.
            """
            normalized_map = {
                normalize_text(col): col
                for col in columns_list
            }

            # 1) Exact normalized candidate match
            for candidate in candidate_names:
                candidate_norm = normalize_text(candidate)
                if candidate_norm in normalized_map:
                    return normalized_map[candidate_norm]

            # 2) Relaxed contains match
            for candidate in candidate_names:
                candidate_norm = normalize_text(candidate)
                for norm_col, original_col in normalized_map.items():
                    if candidate_norm and candidate_norm in norm_col:
                        return original_col

            return None

        contract_col_candidates = [
            "PolicyNo.",
            "Policy No",
            "Policy Number",
            "Policy",
            "Contract Number",
            "Contract No",
            "Contract",
            "Member Number",
            "Member No",
            "Membership Number",
            "Membership No",
            "Account Number",
            "Account No",
        ]

        total_commission_col_candidates = [
            "TotalIncVat",
            "Commission Including VAT",
            "Commission",
            "Commission Amount",
            "Total Commission",
            "Total Commission Payable",
            "Net Commission",
            "Total Due",
        ]

        planner_col_candidates = [
            "Adviser",
            "Advisor",
            "Planner",
            "Broker",
            "Broker Name",
            "Agency",
            "Agency Name",
            "Intermediary",
            "Intermediary Name",
            "Company",
            "Company Name",
            "Practice",
            "Practice Name",
        ]

        client_name_col_candidates = [
            "LifeInsured",
            "Policyholder",
            "Client Name",
            "Client",
            "Customer Name",
            "Customer",
            "Member Name",
            "Member",
            "Insured Name",
            "Name",
        ]

        contract_col = f5_2_1_find_matching_column(list(df.columns), contract_col_candidates)
        total_commission_col = f5_2_1_find_matching_column(list(df.columns), total_commission_col_candidates)
        planner_col = f5_2_1_find_matching_column(list(df.columns), planner_col_candidates)
        client_name_col = f5_2_1_find_matching_column(list(df.columns), client_name_col_candidates)

        print(f"   contract_col        : {contract_col}")
        print(f"   total_commission_col: {total_commission_col}")
        print(f"   planner_col         : {planner_col}")
        print(f"   client_name_col     : {client_name_col}")

        if contract_col is None:
            print("🔴 Could not identify a contract number column in the Bidvest file.")
            return None

        if total_commission_col is None:
            print("🔴 Could not identify a total commission column in the Bidvest file.")
            return None

        # -------------------------------------------------------------------
        # Step 5: Build normalized DataFrame
        # -------------------------------------------------------------------
        print("🟦 Step 5: Build normalized output DataFrame")

        normalized_df = pd.DataFrame()

        # 5.1 Contract Number
        normalized_df["contract_number"] = (
            df[contract_col]
            .apply(normalize_text_preserve_case)
            .astype(str)
            .str.strip()
        )

        # 5.2 Total Commission
        total_series = (
            df[total_commission_col]
            .astype(str)
            .str.replace(",", "", regex=False)
            .str.replace("R", "", regex=False)
            .str.strip()
        )
        normalized_df["total_commission"] = pd.to_numeric(total_series, errors="coerce")

        # 5.3 Client Name
        if client_name_col is not None:
            normalized_df["client_name"] = (
                df[client_name_col]
                .apply(normalize_text_preserve_case)
                .replace("", "tbc")
            )
        else:
            normalized_df["client_name"] = "tbc"

        # 5.4 Planner
        if planner_col is not None:
            normalized_df["planner"] = (
                df[planner_col]
                .apply(normalize_text_preserve_case)
                .replace("", "UNIFY (PTY) LTD")
            )
        else:
            normalized_df["planner"] = "UNIFY (PTY) LTD"

        # 5.5 Fixed standardized fields
        normalized_df["product_house"] = "bidvest"
        normalized_df["commission_month"] = comm_month

        print(f"   normalized_df shape before row filters: {normalized_df.shape}")
        print(f"   normalized_df columns: {list(normalized_df.columns)}")

        # -------------------------------------------------------------------
        # Step 6: Filter invalid rows
        # -------------------------------------------------------------------
        print("🟦 Step 6: Filter invalid normalized rows")

        normalized_df = normalized_df[
            normalized_df["contract_number"].apply(lambda x: normalize_text(x) != "")
        ].copy()

        normalized_df = normalized_df[
            normalized_df["total_commission"].notna()
        ].copy()

        normalized_df = normalized_df.reset_index(drop=True)

        print(f"   normalized_df shape after row filters: {normalized_df.shape}")

        if normalized_df.empty:
            print("🔴 No usable rows remain after normalized filtering.")
            return None

        # -------------------------------------------------------------------
        # Step 7: Reorder final columns
        # -------------------------------------------------------------------
        print("🟦 Step 7: Reorder final columns")

        normalized_df = normalized_df[
            [
                "client_name",
                "product_house",
                "commission_month",
                "contract_number",
                "total_commission",
                "planner",
            ]
        ].copy()

        print("✅ Bidvest normalized read complete")
        print("   final shape:", normalized_df.shape)
        print("   final preview:")
        print(normalized_df.head(10))

        return normalized_df

    except Exception as e:
        print(f"🔴 Error in read_bidvest_df: {e}")
        print(traceback.format_exc())
        return None

# ---------------------------------------------------------------------------
# 5.2.2 Reader: Capital_Legacy
# ---------------------------------------------------------------------------
@register_reader("Capital_Legacy")
def read_capital_legacy_df(file_path, **kwargs):
    """
    Reader for Capital_Legacy commission statement files.
    """
    try:
        print("------------------------------------------------------------")
        print("🧾 5.2.2 Reader: read_capital_legacy_df")
        print(f"   file_path: {file_path}")

        # -------------------------------------------------------------------
        # Step 0: Get passed variables
        # -------------------------------------------------------------------
        print("🟦 Step 0: Get passed variables")
        comm_month = kwargs.get("comm_month")
        print(f"   comm_month: {comm_month}")

        # -------------------------------------------------------------------
        # Step 1: Read raw first sheet with header row
        # -------------------------------------------------------------------
        print("🟦 Step 1: Read raw first sheet with header row")
        df = read_raw_first_sheet(file_path, header=0)
        print(f"   raw shape: {df.shape}")

        if df is None or df.empty:
            print("🔴 Raw Capital_Legacy sheet returned no data.")
            return None

        # -------------------------------------------------------------------
        # Step 2: Normalize headers
        # -------------------------------------------------------------------
        print("🟦 Step 2: Normalize column headers")
        df = normalize_column_headers(df)
        print(f"   normalized columns: {list(df.columns)}")

        # -------------------------------------------------------------------
        # Step 3: Drop blank rows
        # -------------------------------------------------------------------
        print("🟦 Step 3: Drop blank rows")
        df = drop_fully_blank_rows(df)
        print(f"   shape after blank-row cleanup: {df.shape}")

        # -------------------------------------------------------------------
        # Step 4: Find source columns
        # -------------------------------------------------------------------
        print("🟦 Step 4: Detect source columns")

        client_name_col = find_matching_column(df.columns, ["Plan Holder"])
        contract_col = find_matching_column(df.columns, ["Plan Number"])
        total_commission_col = find_matching_column(df.columns, ["Value"])
        planner_col = find_matching_column(df.columns, ["Intermediary"])

        print(f"   client_name_col       : {client_name_col}")
        print(f"   contract_col          : {contract_col}")
        print(f"   total_commission_col  : {total_commission_col}")
        print(f"   planner_col           : {planner_col}")

        if contract_col is None or total_commission_col is None:
            print("🔴 Required Capital_Legacy columns could not be identified.")
            return None

        # -------------------------------------------------------------------
        # Step 5: Build normalized output df
        # -------------------------------------------------------------------
        print("🟦 Step 5: Build normalized output df")
        normalized_df = pd.DataFrame()

        if client_name_col:
            normalized_df["client_name"] = df[client_name_col].apply(normalize_text_preserve_case)
        else:
            normalized_df["client_name"] = ""

        normalized_df["contract_number"] = df[contract_col].apply(normalize_text_preserve_case)
        normalized_df["total_commission"] = coerce_series_to_numeric(df[total_commission_col])

        if planner_col:
            normalized_df["planner"] = (
                df[planner_col]
                .apply(normalize_text_preserve_case)
                .replace("", "UNIFY (PTY) LTD")
            )
        else:
            normalized_df["planner"] = "UNIFY (PTY) LTD"

        normalized_df["product_house"] = "capital_legacy"
        normalized_df["commission_month"] = comm_month

        # -------------------------------------------------------------------
        # Step 6: Filter invalid rows
        # -------------------------------------------------------------------
        print("🟦 Step 6: Filter invalid rows")
        normalized_df = normalized_df[
            normalized_df["contract_number"].apply(lambda x: normalize_text(x) != "")
        ].copy()

        normalized_df = normalized_df[
            normalized_df["total_commission"].notna()
        ].copy()

        normalized_df = normalized_df.reset_index(drop=True)
        print(f"   normalized_df shape after filters: {normalized_df.shape}")

        if normalized_df.empty:
            print("🔴 No usable Capital_Legacy rows remain after filtering.")
            return None

        # -------------------------------------------------------------------
        # Step 7: Reorder final columns
        # -------------------------------------------------------------------
        print("🟦 Step 7: Reorder final columns")
        normalized_df = normalized_df[
            [
                "client_name",
                "product_house",
                "commission_month",
                "contract_number",
                "total_commission",
                "planner",
            ]
        ].copy()

        print("✅ Capital_Legacy normalized read complete")
        print(normalized_df.head(10))

        return normalized_df

    except Exception as e:
        print(f"🔴 Error in read_capital_legacy_df: {e}")
        print(traceback.format_exc())
        return None

# ---------------------------------------------------------------------------
# 5.2.3 Reader: Cura
# ---------------------------------------------------------------------------
@register_reader("Cura")
def read_cura_df(file_path, **kwargs):
    """
    Reader for Cura commission statement files.


    Logic:
    - read first sheet using first row as header
    - normalize headers
    - drop fully blank rows
    - map the source columns into the normalized structure
    """
    try:
        print("------------------------------------------------------------")
        print("🧾 5.2.3 Reader: read_cura_df")
        print(f"   file_path: {file_path}")


        # -------------------------------------------------------------------
        # Step 0: Get Variables
        # -------------------------------------------------------------------
        print("🟦 Step 0: Get passed variables")
        comm_month = kwargs.get("comm_month")
        comm_tables_main_df = kwargs.get("comm_tables_main_df")


        print(f"   comm_month: {comm_month}")
        print(
            f"   comm_tables_main_df shape: "
            f"{comm_tables_main_df.shape if isinstance(comm_tables_main_df, pd.DataFrame) else None}"
        )


        # -------------------------------------------------------------------
        # Step 1: Read raw first sheet with header row
        # -------------------------------------------------------------------
        print("🟦 Step 1: Read raw first sheet with header row")
        df = read_raw_first_sheet(file_path, header=0)
        print(f"   raw shape: {df.shape}")
        print(f"   raw columns before normalization: {list(df.columns)}")


        if df is None or df.empty:
            print("🔴 Raw Cura sheet returned no data.")
            return None


        # -------------------------------------------------------------------
        # Step 2: Normalize headers
        # -------------------------------------------------------------------
        print("🟦 Step 2: Normalize column headers")
        df = normalize_column_headers(df)
        print(f"   normalized columns: {list(df.columns)}")


        # -------------------------------------------------------------------
        # Step 3: Drop fully blank rows
        # -------------------------------------------------------------------
        print("🟦 Step 3: Drop fully blank rows")
        df = drop_fully_blank_rows(df)
        print(f"   shape after dropping blank rows: {df.shape}")


        if df.empty:
            print("🔴 No usable rows remain after blank-row cleanup.")
            return None


        # -------------------------------------------------------------------
        # Step 4: Detect source columns
        # -------------------------------------------------------------------
        print("🟦 Step 4: Detect source columns for normalized mapping")


        def f5_2_2_find_matching_column(columns_list, candidate_names):
            """
            Match a real source column by comparing normalized text.
            """
            normalized_map = {
                normalize_text(col): col
                for col in columns_list
            }


            # 1) Exact normalized candidate match
            for candidate in candidate_names:
                candidate_norm = normalize_text(candidate)
                if candidate_norm in normalized_map:
                    return normalized_map[candidate_norm]


            # 2) Relaxed contains match
            for candidate in candidate_names:
                candidate_norm = normalize_text(candidate)
                for norm_col, original_col in normalized_map.items():
                    if candidate_norm and candidate_norm in norm_col:
                        return original_col


            return None


        contract_col_candidates = [
            "Policy #",
            "Policy Number",
            "Policy No",
            "Contract Number",
            "Contract No",
            "Member Number",
            "Account Number",
        ]


        total_commission_col_candidates = [
            "Commission",
            "Commission Amount",
            "Total Commission",
            "Net Commission",
        ]


        planner_col_candidates = [
            "Agent",
            "Planner",
            "Broker",
            "Broker Name",
            "Introducer",
        ]


        client_name_col_candidates = [
            "Client",
            "Client Name",
            "Member Name",
            "Policyholder",
            "Name",
        ]


        contract_col = f5_2_2_find_matching_column(list(df.columns), contract_col_candidates)
        total_commission_col = f5_2_2_find_matching_column(list(df.columns), total_commission_col_candidates)
        planner_col = f5_2_2_find_matching_column(list(df.columns), planner_col_candidates)
        client_name_col = f5_2_2_find_matching_column(list(df.columns), client_name_col_candidates)


        print(f"   contract_col        : {contract_col}")
        print(f"   total_commission_col: {total_commission_col}")
        print(f"   planner_col         : {planner_col}")
        print(f"   client_name_col     : {client_name_col}")


        if contract_col is None:
            print("🔴 Could not identify a contract number column in the Cura file.")
            return None


        if total_commission_col is None:
            print("🔴 Could not identify a total commission column in the Cura file.")
            return None


        # -------------------------------------------------------------------
        # Step 5: Build normalized DataFrame
        # -------------------------------------------------------------------
        print("🟦 Step 5: Build normalized output DataFrame")


        normalized_df = pd.DataFrame()


        # 5.1 Contract Number
        normalized_df["contract_number"] = (
            df[contract_col]
            .apply(normalize_text_preserve_case)
            .astype(str)
            .str.strip()
        )


        # 5.2 Total Commission
        total_series = (
            df[total_commission_col]
            .astype(str)
            .str.replace(",", "", regex=False)
            .str.replace("R", "", regex=False)
            .str.strip()
        )
        normalized_df["total_commission"] = pd.to_numeric(total_series, errors="coerce")


        # 5.3 Client Name
        if client_name_col is not None:
            normalized_df["client_name"] = (
                df[client_name_col]
                .apply(normalize_text_preserve_case)
                .replace("", "tbc")
            )
        else:
            normalized_df["client_name"] = "tbc"


        # 5.4 Planner
        if planner_col is not None:
            normalized_df["planner"] = (
                df[planner_col]
                .apply(normalize_text_preserve_case)
                .replace("", "tbc")
            )
        else:
            normalized_df["planner"] = "tbc"


        # 5.5 Fixed standardized fields
        normalized_df["product_house"] = "cura"
        normalized_df["commission_month"] = comm_month


        print(f"   normalized_df shape before row filters: {normalized_df.shape}")
        print(f"   normalized_df columns: {list(normalized_df.columns)}")


        # -------------------------------------------------------------------
        # Step 6: Filter invalid rows
        # -------------------------------------------------------------------
        print("🟦 Step 6: Filter invalid normalized rows")


        normalized_df = normalized_df[
            normalized_df["contract_number"].apply(lambda x: normalize_text(x) != "")
        ].copy()


        normalized_df = normalized_df[
            normalized_df["total_commission"].notna()
        ].copy()


        normalized_df = normalized_df.reset_index(drop=True)


        print(f"   normalized_df shape after row filters: {normalized_df.shape}")


        # -------------------------------------------------------------------
        # Step 7: Reorder final columns
        # -------------------------------------------------------------------
        print("🟦 Step 7: Reorder final columns")


        normalized_df = normalized_df[
            [
                "client_name",
                "product_house",
                "commission_month",
                "contract_number",
                "total_commission",
                "planner",
            ]
        ].copy()


        print("✅ Cura normalized read complete")
        print("   final shape:", normalized_df.shape)
        print("   final preview:")
        print(normalized_df.head(10))


        return normalized_df


    except Exception as e:
        print(f"🔴 Error in read_cura_df: {e}")
        print(traceback.format_exc())
        return None

# ---------------------------------------------------------------------------
# 5.2.4 Reader: Bonitas
# ---------------------------------------------------------------------------
@register_reader("Bonitas")
def read_bonitas_df(file_path, **kwargs):
    """
    Reader for Bonitas commission statement files.

    Logic:
    - read the best matching sheet in the workbook
    - normalize headers
    - remove total / subtotal rows
    - map source columns into the normalized structure
    """
    try:
        print("------------------------------------------------------------")
        print("🧾 5.2.4 Reader: read_bonitas_df")
        print(f"   file_path: {file_path}")

        # -------------------------------------------------------------------
        # Step 0: Get Variables
        # -------------------------------------------------------------------
        print("🟦 Step 0: Get passed variables")
        comm_month = kwargs.get("comm_month")
        comm_tables_main_df = kwargs.get("comm_tables_main_df")

        print(f"   comm_month: {comm_month}")
        print(
            f"   comm_tables_main_df shape: "
            f"{comm_tables_main_df.shape if isinstance(comm_tables_main_df, pd.DataFrame) else None}"
        )

        # -------------------------------------------------------------------
        # Step 1: Read best matching sheet
        # -------------------------------------------------------------------
        print("🟦 Step 1: Read best matching Bonitas sheet")

        required_column_candidates = [
            ["Membership Number", "External Member Number"],
            ["Commission Amount (Excluding VAT)", "Amount to be Paid"],
        ]

        source_sheet_name, df = read_best_matching_sheet_by_required_columns(
            file_path=file_path,
            required_column_candidates=required_column_candidates,
            header=3,
            exclude_sheet_names=["processed_data"]
        )

        print(f"   source_sheet_name: {source_sheet_name}")

        if df is None or df.empty:
            print("🔴 No usable Bonitas sheet could be identified.")
            return None

        print(f"   raw shape: {df.shape}")
        print(f"   columns after normalization: {list(df.columns)}")

        # -------------------------------------------------------------------
        # Step 2: Drop blank rows
        # -------------------------------------------------------------------
        print("🟦 Step 2: Drop blank rows")
        df = drop_fully_blank_rows(df)
        print(f"   shape after blank-row cleanup: {df.shape}")

        if df.empty:
            print("🔴 No usable Bonitas rows remain after blank-row cleanup.")
            return None

        # -------------------------------------------------------------------
        # Step 3: Detect source columns
        # -------------------------------------------------------------------
        print("🟦 Step 3: Detect source columns")

        planner_col = find_matching_column(df.columns, ["Name", "Broker Name"])
        contract_col = find_matching_column(df.columns, ["Membership Number", "External Member Number"])
        title_col = find_matching_column(df.columns, ["Title"])
        initials_col = find_matching_column(df.columns, ["Initials"])
        surname_col = find_matching_column(df.columns, ["Surname"])
        total_commission_col = find_matching_column(
            df.columns,
            ["Commission Amount (Excluding VAT)", "Amount to be Paid"]
        )
        reason_col = find_matching_column(df.columns, ["Reason"])

        first_col = df.columns[0] if len(df.columns) > 0 else None

        print(f"   planner_col           : {planner_col}")
        print(f"   contract_col          : {contract_col}")
        print(f"   title_col             : {title_col}")
        print(f"   initials_col          : {initials_col}")
        print(f"   surname_col           : {surname_col}")
        print(f"   total_commission_col  : {total_commission_col}")
        print(f"   reason_col            : {reason_col}")
        print(f"   first_col             : {first_col}")

        if contract_col is None or total_commission_col is None:
            print("🔴 Required Bonitas columns could not be identified.")
            return None

        # -------------------------------------------------------------------
        # Step 4: Remove subtotal / total rows
        # -------------------------------------------------------------------
        print("🟦 Step 4: Remove total / subtotal rows")

        if first_col is not None:
            total_row_mask = (
                df[first_col]
                .astype(str)
                .str.strip()
                .str.lower()
                .str.startswith("total for")
            )
        else:
            total_row_mask = pd.Series(False, index=df.index)

        print(f"   total_row_count: {int(total_row_mask.sum())}")

        df = df[~total_row_mask].copy().reset_index(drop=True)
        print(f"   shape after removing total rows: {df.shape}")

        # -------------------------------------------------------------------
        # Step 5: Build normalized output df
        # -------------------------------------------------------------------
        print("🟦 Step 5: Build normalized output DataFrame")

        normalized_df = pd.DataFrame(index=df.index)

        normalized_df["contract_number"] = (
            df[contract_col]
            .apply(normalize_text_preserve_case)
            .astype(str)
            .str.strip()
        )

        normalized_df["total_commission"] = coerce_series_to_numeric(df[total_commission_col])

        title_series = df[title_col].apply(normalize_text_preserve_case) if title_col else pd.Series("", index=df.index)
        initials_series = df[initials_col].apply(normalize_text_preserve_case) if initials_col else pd.Series("", index=df.index)
        surname_series = df[surname_col].apply(normalize_text_preserve_case) if surname_col else pd.Series("", index=df.index)

        if title_col and initials_col and surname_col:
            normalized_df["client_name"] = (
                title_series.astype(str).str.strip()
                + " "
                + initials_series.astype(str).str.strip()
                + " "
                + surname_series.astype(str).str.strip()
            ).str.replace(r"\s+", " ", regex=True).str.strip()
        elif surname_col:
            normalized_df["client_name"] = surname_series.astype(str).str.strip()
        else:
            normalized_df["client_name"] = ""

        if planner_col:
            normalized_df["planner"] = (
                df[planner_col]
                .apply(normalize_text_preserve_case)
                .replace("", "UNIFY (PTY) LTD")
            )
        else:
            normalized_df["planner"] = "UNIFY (PTY) LTD"

        normalized_df["product_house"] = "bonitas"
        normalized_df["commission_month"] = comm_month

        print(f"   normalized_df shape before filters: {normalized_df.shape}")

        # -------------------------------------------------------------------
        # Step 6: Filter invalid rows
        # -------------------------------------------------------------------
        print("🟦 Step 6: Filter invalid rows")

        normalized_df = normalized_df[
            normalized_df["contract_number"].apply(lambda x: normalize_text(x) != "")
        ].copy()

        normalized_df = normalized_df[
            normalized_df["total_commission"].notna()
        ].copy()

        if reason_col is not None:
            reason_series = df.loc[normalized_df.index, reason_col]
            normalized_df = normalized_df[
                reason_series.apply(lambda x: normalize_text(x) != "")
            ].copy()

        normalized_df = normalized_df.reset_index(drop=True)
        print(f"   normalized_df shape after filters: {normalized_df.shape}")

        if normalized_df.empty:
            print("🔴 No usable Bonitas rows remain after filtering.")
            return None

        # -------------------------------------------------------------------
        # Step 7: Reorder final columns
        # -------------------------------------------------------------------
        print("🟦 Step 7: Reorder final columns")

        normalized_df = normalized_df[
            [
                "client_name",
                "product_house",
                "commission_month",
                "contract_number",
                "total_commission",
                "planner",
            ]
        ].copy()

        print("✅ Bonitas normalized read complete")
        print("   final shape:", normalized_df.shape)
        print("   final preview:")
        print(normalized_df.head(10))

        return normalized_df

    except Exception as e:
        print(f"🔴 Error in read_bonitas_df: {e}")
        print(traceback.format_exc())
        return None
    
# ---------------------------------------------------------------------------
# 5.2.5 Reader: Hollard Life
# ---------------------------------------------------------------------------
@register_reader("Hollard_Life")
def read_hollard_life_df(file_path, **kwargs):
    """
    Reader for Hollard Life commission statement files.

    Goal:
    - read the raw Hollard Life sheet correctly
    - normalize the source headers
    - clean the raw data
    - detect the main source columns robustly
    - map the source columns into the normalized structure

    Returns normalized columns:
    - client_name
    - product_house
    - commission_month
    - contract_number
    - total_commission
    - planner
    """
    try:
        print("------------------------------------------------------------")
        print("🧾 5.2.x Reader: read_hollard_life_df")
        print(f"   file_path: {file_path}")

        # -------------------------------------------------------------------
        # Step 0: Get passed variables
        # -------------------------------------------------------------------
        print("🟦 Step 0: Get passed variables")
        comm_month = kwargs.get("comm_month")
        comm_tables_main_df = kwargs.get("comm_tables_main_df")

        print(f"   comm_month: {comm_month}")
        print(
            f"   comm_tables_main_df shape: "
            f"{comm_tables_main_df.shape if isinstance(comm_tables_main_df, pd.DataFrame) else None}"
        )

        # -------------------------------------------------------------------
        # Step 1: Read raw first sheet with header row
        # -------------------------------------------------------------------
        print("🟦 Step 1: Read raw first sheet with header row")
        df = read_raw_first_sheet(file_path, header=0)
        print(f"   raw shape: {df.shape}")
        print(f"   raw columns before normalization: {list(df.columns)}")

        if df is None or df.empty:
            print("🔴 Raw Hollard Life sheet returned no data.")
            return None

        # -------------------------------------------------------------------
        # Step 2: Normalize column headers
        # -------------------------------------------------------------------
        print("🟦 Step 2: Normalize column headers")
        df = normalize_column_headers(df)
        print(f"   normalized columns: {list(df.columns)}")

        # -------------------------------------------------------------------
        # Step 3: Drop fully blank rows
        # -------------------------------------------------------------------
        print("🟦 Step 3: Drop fully blank rows")
        df = drop_fully_blank_rows(df)
        print(f"   shape after blank-row cleanup: {df.shape}")

        if df.empty:
            print("🔴 No usable Hollard Life rows remain after blank-row cleanup.")
            return None

        # -------------------------------------------------------------------
        # Step 4: Detect source columns
        # -------------------------------------------------------------------
        print("🟦 Step 4: Detect source columns")

        client_name_col = find_matching_column(
            df.columns,
            ["Policyholder", "Policy Holder", "Client Name", "Client"]
        )
        contract_col = find_matching_column(
            df.columns,
            ["Policy no.", "Policy no", "Policy Number", "Policy Number"]
        )
        total_commission_col = find_matching_column(
            df.columns,
            ["Amount", "Commission", "Commission Amount"]
        )
        planner_col = find_matching_column(
            df.columns,
            ["Intermediary", "Broker", "Planner"]
        )

        print(f"   client_name_col       : {client_name_col}")
        print(f"   contract_col          : {contract_col}")
        print(f"   total_commission_col  : {total_commission_col}")
        print(f"   planner_col           : {planner_col}")

        if contract_col is None or total_commission_col is None:
            print("🔴 Required Hollard Life columns could not be identified.")
            return None

        # -------------------------------------------------------------------
        # Step 5: Build normalized output df
        # -------------------------------------------------------------------
        print("🟦 Step 5: Build normalized output df")
        normalized_df = pd.DataFrame()

        if client_name_col:
            normalized_df["client_name"] = (
                df[client_name_col]
                .apply(normalize_text_preserve_case)
                .replace("", "tbc")
            )
        else:
            normalized_df["client_name"] = "tbc"

        normalized_df["contract_number"] = (
            df[contract_col]
            .apply(normalize_text_preserve_case)
        )

        normalized_df["total_commission"] = coerce_series_to_numeric(
            df[total_commission_col]
        )

        if planner_col:
            normalized_df["planner"] = (
                df[planner_col]
                .apply(normalize_text_preserve_case)
                .replace("", "UNIFY (PTY) LTD")
            )
        else:
            normalized_df["planner"] = "UNIFY (PTY) LTD"

        normalized_df["product_house"] = "hollard_life"
        normalized_df["commission_month"] = comm_month

        # -------------------------------------------------------------------
        # Step 6: Filter invalid rows
        # -------------------------------------------------------------------
        print("🟦 Step 6: Filter invalid rows")
        normalized_df = normalized_df[
            normalized_df["contract_number"].apply(lambda x: normalize_text(x) != "")
        ].copy()

        normalized_df = normalized_df[
            normalized_df["total_commission"].notna()
        ].copy()

        normalized_df = normalized_df.reset_index(drop=True)
        print(f"   normalized_df shape after filters: {normalized_df.shape}")

        if normalized_df.empty:
            print("🔴 No usable Hollard Life rows remain after filtering.")
            return None

        # -------------------------------------------------------------------
        # Step 7: Reorder final columns
        # -------------------------------------------------------------------
        print("🟦 Step 7: Reorder final columns")
        normalized_df = normalized_df[
            [
                "client_name",
                "product_house",
                "commission_month",
                "contract_number",
                "total_commission",
                "planner",
            ]
        ].copy()

        print("✅ Hollard Life normalized read complete")
        print(normalized_df.head(10))

        return normalized_df

    except Exception as e:
        print(f"🔴 Error in read_hollard_life_df: {e}")
        print(traceback.format_exc())
        return None

# ---------------------------------------------------------------------------
# 5.2.6 Reader: Liberty
# ---------------------------------------------------------------------------
@register_reader("Liberty")
def read_liberty_df(file_path, **kwargs):
    """
    Reader for Liberty commission statement files.

    Goal:
    - read the raw Liberty sheet correctly
    - normalize the source headers
    - clean the raw data
    - detect the main source columns robustly
    - map the source columns into the normalized structure

    Returns normalized columns:
    - client_name
    - product_house
    - commission_month
    - contract_number
    - total_commission
    - planner
    """
    try:
        print("------------------------------------------------------------")
        print("🧾 5.2.x Reader: read_liberty_df")
        print(f"   file_path: {file_path}")

        # -------------------------------------------------------------------
        # Step 0: Get passed variables
        # -------------------------------------------------------------------
        print("🟦 Step 0: Get passed variables")
        comm_month = kwargs.get("comm_month")
        comm_tables_main_df = kwargs.get("comm_tables_main_df")

        print(f"   comm_month: {comm_month}")
        print(
            f"   comm_tables_main_df shape: "
            f"{comm_tables_main_df.shape if isinstance(comm_tables_main_df, pd.DataFrame) else None}"
        )

        # -------------------------------------------------------------------
        # Step 1: Read raw first sheet with header row
        # -------------------------------------------------------------------
        print("🟦 Step 1: Read raw first sheet with header row")
        df = read_raw_first_sheet(file_path, header=0)
        print(f"   raw shape: {df.shape}")
        print(f"   raw columns before normalization: {list(df.columns)}")

        if df is None or df.empty:
            print("🔴 Raw Liberty sheet returned no data.")
            return None

        # -------------------------------------------------------------------
        # Step 2: Normalize column headers
        # -------------------------------------------------------------------
        print("🟦 Step 2: Normalize column headers")
        df = normalize_column_headers(df)
        print(f"   normalized columns: {list(df.columns)}")

        # -------------------------------------------------------------------
        # Step 3: Drop fully blank rows
        # -------------------------------------------------------------------
        print("🟦 Step 3: Drop fully blank rows")
        df = drop_fully_blank_rows(df)
        print(f"   shape after blank-row cleanup: {df.shape}")

        if df.empty:
            print("🔴 No usable Liberty rows remain after blank-row cleanup.")
            return None

        # -------------------------------------------------------------------
        # Step 4: Detect source columns
        # -------------------------------------------------------------------
        print("🟦 Step 4: Detect source columns")

        client_name_col = find_matching_column(
            df.columns,
            ["Life Assured", "Client Name", "Client", "Policyholder"]
        )
        contract_col = find_matching_column(
            df.columns,
            ["Contract", "Contract Number", "Policy Number"]
        )
        total_commission_col = find_matching_column(
            df.columns,
            ["Amount", "Commission", "Commission Amount"]
        )
        planner_col = find_matching_column(
            df.columns,
            ["Planner", "Intermediary", "Broker"]
        )

        print(f"   client_name_col       : {client_name_col}")
        print(f"   contract_col          : {contract_col}")
        print(f"   total_commission_col  : {total_commission_col}")
        print(f"   planner_col           : {planner_col}")

        if contract_col is None or total_commission_col is None:
            print("🔴 Required Liberty columns could not be identified.")
            return None

        # -------------------------------------------------------------------
        # Step 5: Build normalized output df
        # -------------------------------------------------------------------
        print("🟦 Step 5: Build normalized output df")
        normalized_df = pd.DataFrame()

        if client_name_col:
            normalized_df["client_name"] = (
                df[client_name_col]
                .apply(normalize_text_preserve_case)
                .replace("", "tbc")
            )
        else:
            normalized_df["client_name"] = "tbc"

        normalized_df["contract_number"] = (
            df[contract_col]
            .apply(normalize_text_preserve_case)
        )

        normalized_df["total_commission"] = coerce_series_to_numeric(
            df[total_commission_col]
        )

        if planner_col:
            normalized_df["planner"] = (
                df[planner_col]
                .apply(normalize_text_preserve_case)
                .replace("", "Unify")
            )
        else:
            normalized_df["planner"] = "Unify"

        normalized_df["product_house"] = "liberty"
        normalized_df["commission_month"] = comm_month

        # -------------------------------------------------------------------
        # Step 6: Filter invalid rows
        # -------------------------------------------------------------------
        print("🟦 Step 6: Filter invalid rows")
        normalized_df = normalized_df[
            normalized_df["contract_number"].apply(lambda x: normalize_text(x) != "")
        ].copy()

        normalized_df = normalized_df[
            normalized_df["total_commission"].notna()
        ].copy()

        normalized_df = normalized_df.reset_index(drop=True)
        print(f"   normalized_df shape after filters: {normalized_df.shape}")

        if normalized_df.empty:
            print("🔴 No usable Liberty rows remain after filtering.")
            return None

        # -------------------------------------------------------------------
        # Step 7: Reorder final columns
        # -------------------------------------------------------------------
        print("🟦 Step 7: Reorder final columns")
        normalized_df = normalized_df[
            [
                "client_name",
                "product_house",
                "commission_month",
                "contract_number",
                "total_commission",
                "planner",
            ]
        ].copy()

        print("✅ Liberty normalized read complete")
        print(normalized_df.head(10))

        return normalized_df

    except Exception as e:
        print(f"🔴 Error in read_liberty_df: {e}")
        print(traceback.format_exc())
        return None

# ---------------------------------------------------------------------------
# 5.1.x Reader: Momentum_MFP
# ---------------------------------------------------------------------------
@register_reader("Momentum_MFP")
def read_momentum_mfp_df(file_path, **kwargs):
    """
    Reader for Momentum MFP commission statement files.


    Note:
    The uploaded sample is a legacy .xls file with the true header row on Excel
    row 8. The total payable amount is best taken from 'Grand Total'.


    Expected output columns:
    - client_name
    - product_house
    - commission_month
    - contract_number
    - total_commission
    - planner
    """
    try:
        print("------------------------------------------------------------")
        print("🧾 5.1.x Reader: read_momentum_mfp_df")
        print(f"   file_path: {file_path}")


        # -------------------------------------------------------------------
        # Step 0: Get passed variables
        # -------------------------------------------------------------------
        print("🟦 Step 0: Get passed variables")
        comm_month = kwargs.get("comm_month")
        print(f"   comm_month: {comm_month}")


        # -------------------------------------------------------------------
        # Step 1: Read raw first sheet using header row 8
        # -------------------------------------------------------------------
        print("🟦 Step 1: Read raw first sheet using header row 8")
        df = read_raw_first_sheet(file_path, header=7)
        print(f"   raw shape: {df.shape}")


        if df is None or df.empty:
            print("🔴 Raw Momentum MFP sheet returned no data.")
            return None


        # -------------------------------------------------------------------
        # Step 2: Normalize headers
        # -------------------------------------------------------------------
        print("🟦 Step 2: Normalize column headers")
        df = normalize_column_headers(df)
        print(f"   normalized columns: {list(df.columns)}")


        # -------------------------------------------------------------------
        # Step 3: Drop blank rows
        # -------------------------------------------------------------------
        print("🟦 Step 3: Drop blank rows")
        df = drop_fully_blank_rows(df)
        print(f"   shape after blank-row cleanup: {df.shape}")


        # -------------------------------------------------------------------
        # Step 4: Detect source columns
        # -------------------------------------------------------------------
        print("🟦 Step 4: Detect source columns")
        client_name_col = find_matching_column(
            df.columns,
            ["Client", "Client Name", "Policyholder"]
        )
        contract_col = find_matching_column(
            df.columns,
            ["Contract number", "Contract No", "Policy Number"]
        )
        total_commission_col = find_matching_column(
            df.columns,
            ["Grand Total", "Financial planner commission", "Franchise House Total"]
        )
        planner_col = find_matching_column(
            df.columns,
            ["Planner", "Broker", "Agent"]
        )


        print(f"   client_name_col       : {client_name_col}")
        print(f"   contract_col          : {contract_col}")
        print(f"   total_commission_col  : {total_commission_col}")
        print(f"   planner_col           : {planner_col}")


        if contract_col is None or total_commission_col is None:
            print("🔴 Required Momentum MFP columns could not be identified.")
            return None


        # -------------------------------------------------------------------
        # Step 5: Build normalized output df
        # -------------------------------------------------------------------
        print("🟦 Step 5: Build normalized output df")
        normalized_df = pd.DataFrame()


        if client_name_col:
            normalized_df["client_name"] = df[client_name_col].apply(normalize_text_preserve_case)
        else:
            normalized_df["client_name"] = ""


        normalized_df["product_house"] = "momentum_mfp"
        normalized_df["commission_month"] = comm_month


        contract_series = df[contract_col]
        normalized_df["contract_number"] = contract_series.apply(
            lambda x: ""
            if pd.isna(x)
            else (
                str(int(x))
                if isinstance(x, (int, float)) and float(x).is_integer()
                else normalize_text_preserve_case(x)
            )
        )


        normalized_df["total_commission"] = coerce_series_to_numeric(df[total_commission_col])


        if planner_col:
            normalized_df["planner"] = df[planner_col].apply(normalize_text_preserve_case)
        else:
            normalized_df["planner"] = ""


        print(f"   normalized_df shape before filters: {normalized_df.shape}")


        # -------------------------------------------------------------------
        # Step 6: Filter invalid rows
        # -------------------------------------------------------------------
        print("🟦 Step 6: Filter invalid rows")
        normalized_df = normalized_df[
            normalized_df["contract_number"].apply(lambda x: normalize_text(x) != "")
        ].copy()


        normalized_df = normalized_df[
            normalized_df["total_commission"].notna()
        ].copy()


        normalized_df = normalized_df[
            ~normalized_df["client_name"].apply(
                lambda x: normalize_text(x) in ["", "total", "sub total", "subtotal", "grand total"]
            )
        ].copy()


        normalized_df = normalized_df.reset_index(drop=True)
        print(f"   normalized_df shape after filters: {normalized_df.shape}")


        if normalized_df.empty:
            print("🔴 No usable Momentum MFP rows remain after filtering.")
            return None


        # -------------------------------------------------------------------
        # Step 7: Reorder final columns
        # -------------------------------------------------------------------
        print("🟦 Step 7: Reorder final columns")
        normalized_df = normalized_df[
            [
                "client_name",
                "product_house",
                "commission_month",
                "contract_number",
                "total_commission",
                "planner",
            ]
        ].copy()


        print("✅ Momentum MFP normalized read complete")
        print(normalized_df.head(10))


        return normalized_df


    except Exception as e:
        print(f"🔴 Error in read_momentum_mfp_df: {e}")
        print(traceback.format_exc())
        return None
    
# ---------------------------------------------------------------------------
# 5.2.x Reader: Old_Mutual_Short_Term
# ---------------------------------------------------------------------------
@register_reader("Old_Mutual_Short_Term")
def read_old_mutual_short_term_df(file_path, **kwargs):
    """
    Reader for Old Mutual Short Term commission statement files.


    Expected output columns:
    - client_name
    - product_house
    - commission_month
    - contract_number
    - total_commission
    - planner
    """
    try:
        print("------------------------------------------------------------")
        print("🧾 5.2.x Reader: read_old_mutual_short_term_df")
        print(f"   file_path: {file_path}")


        # -------------------------------------------------------------------
        # Step 0: Get passed variables
        # -------------------------------------------------------------------
        print("🟦 Step 0: Get passed variables")
        comm_month = kwargs.get("comm_month")
        print(f"   comm_month: {comm_month}")


        # -------------------------------------------------------------------
        # Step 1: Read raw first sheet with row 1 as header
        # -------------------------------------------------------------------
        print("🟦 Step 1: Read raw first sheet with header row")
        df = read_raw_first_sheet(file_path, header=0)
        print(f"   raw shape: {df.shape}")


        if df is None or df.empty:
            print("🔴 Raw Old Mutual Short Term sheet returned no data.")
            return None


        # -------------------------------------------------------------------
        # Step 2: Normalize headers
        # -------------------------------------------------------------------
        print("🟦 Step 2: Normalize column headers")
        df = normalize_column_headers(df)
        print(f"   normalized columns: {list(df.columns)}")


        # -------------------------------------------------------------------
        # Step 3: Drop blank rows
        # -------------------------------------------------------------------
        print("🟦 Step 3: Drop fully blank rows")
        df = drop_fully_blank_rows(df)
        print(f"   shape after blank-row cleanup: {df.shape}")


        # -------------------------------------------------------------------
        # Step 4: Detect source columns
        # -------------------------------------------------------------------
        print("🟦 Step 4: Detect source columns")
        client_name_col = find_matching_column(
            df.columns,
            ["Insured Name", "Insured Name Section Code Split", "Client Name", "Client"]
        )
        contract_col = find_matching_column(
            df.columns,
            ["Policy Number", "Policy No", "Policy #", "Contract Number"]
        )
        total_commission_col = find_matching_column(
            df.columns,
            ["*Commission", "Commission", "Commission Amount"]
        )
        planner_col = find_matching_column(
            df.columns,
            ["Planner", "Broker", "Agent"]
        )


        print(f"   client_name_col       : {client_name_col}")
        print(f"   contract_col          : {contract_col}")
        print(f"   total_commission_col  : {total_commission_col}")
        print(f"   planner_col           : {planner_col}")


        if contract_col is None or total_commission_col is None:
            print("🔴 Required Old Mutual Short Term columns could not be identified.")
            return None


        # -------------------------------------------------------------------
        # Step 5: Build normalized output df
        # -------------------------------------------------------------------
        print("🟦 Step 5: Build normalized output df")
        normalized_df = pd.DataFrame()


        if client_name_col:
            normalized_df["client_name"] = df[client_name_col].apply(normalize_text_preserve_case)
        else:
            normalized_df["client_name"] = ""


        normalized_df["product_house"] = "old_mutual_short_term"
        normalized_df["commission_month"] = comm_month


        contract_series = df[contract_col]
        normalized_df["contract_number"] = contract_series.apply(
            lambda x: ""
            if pd.isna(x)
            else (
                str(int(x))
                if isinstance(x, (int, float)) and float(x).is_integer()
                else normalize_text_preserve_case(x)
            )
        )


        normalized_df["total_commission"] = coerce_series_to_numeric(df[total_commission_col])


        if planner_col:
            normalized_df["planner"] = df[planner_col].apply(normalize_text_preserve_case)
        else:
            normalized_df["planner"] = ""


        # -------------------------------------------------------------------
        # Step 6: Filter invalid rows
        # -------------------------------------------------------------------
        print("🟦 Step 6: Filter invalid rows")
        normalized_df = normalized_df[
            normalized_df["contract_number"].apply(lambda x: normalize_text(x) != "")
        ].copy()


        normalized_df = normalized_df[
            normalized_df["total_commission"].notna()
        ].copy()


        normalized_df = normalized_df.reset_index(drop=True)
        print(f"   normalized_df shape after filters: {normalized_df.shape}")


        if normalized_df.empty:
            print("🔴 No usable Old Mutual Short Term rows remain after filtering.")
            return None


        # -------------------------------------------------------------------
        # Step 7: Reorder final columns
        # -------------------------------------------------------------------
        print("🟦 Step 7: Reorder final columns")
        normalized_df = normalized_df[
            [
                "client_name",
                "product_house",
                "commission_month",
                "contract_number",
                "total_commission",
                "planner",
            ]
        ].copy()


        print("✅ Old Mutual Short Term normalized read complete")
        print(normalized_df.head(10))


        return normalized_df


    except Exception as e:
        print(f"🔴 Error in read_old_mutual_short_term_df: {e}")
        print(traceback.format_exc())
        return None

# ---------------------------------------------------------------------------
# 5.2.x Reader: Old_Mutual_Life_Invest
# ---------------------------------------------------------------------------
@register_reader("Old_Mutual_Life_Invest")
def read_old_mutual_life_invest_df(file_path, **kwargs):
    """
    Reader for Old Mutual Life / Invest commission statement files.


    Practical note:
    - the uploaded sample currently only contains a processed_data sheet
    - so this reader first checks whether the workbook is already in a
      normalized/processed layout
    - if not, it then attempts a broad raw-sheet read using flexible
      fixed-header matching


    Expected output columns:
    - client_name
    - product_house
    - commission_month
    - contract_number
    - total_commission
    - planner
    """
    try:
        print("------------------------------------------------------------")
        print("🧾 5.2.x Reader: read_old_mutual_life_invest_df")
        print(f"   file_path: {file_path}")


        # -------------------------------------------------------------------
        # Step 0: Get passed variables
        # -------------------------------------------------------------------
        print("🟦 Step 0: Get passed variables")
        comm_month = kwargs.get("comm_month")
        print(f"   comm_month: {comm_month}")


        # -------------------------------------------------------------------
        # Step 1: Read best matching sheet
        # -------------------------------------------------------------------
        print("🟦 Step 1: Read best matching Old Mutual Life Invest sheet")


        required_column_candidates = [
            ["Client Name", "Policyholder", "Investor Name", "Insured Name", "Name"],
            ["Contract Number", "Policy Number", "Policy No", "Plan Number", "Account Number"],
            ["Total Commission", "Commission", "Commission Amount", "Amount", "Net Commission"],
        ]


        source_sheet_name, df = read_best_matching_sheet_by_required_columns(
            file_path=file_path,
            required_column_candidates=required_column_candidates,
            header=0,
            exclude_sheet_names=[]
        )


        print(f"   source_sheet_name: {source_sheet_name}")


        if df is None or df.empty:
            print("🔴 No usable Old Mutual Life Invest sheet could be identified.")
            return None


        print(f"   raw shape: {df.shape}")
        print(f"   columns after normalization: {list(df.columns)}")


        # -------------------------------------------------------------------
        # Step 2: Drop blank rows
        # -------------------------------------------------------------------
        print("🟦 Step 2: Drop blank rows")
        df = drop_fully_blank_rows(df)
        print(f"   shape after blank-row cleanup: {df.shape}")


        if df.empty:
            print("🔴 No usable Old Mutual Life Invest rows remain after blank-row cleanup.")
            return None


        # -------------------------------------------------------------------
        # Step 3: Detect source columns
        # -------------------------------------------------------------------
        print("🟦 Step 3: Detect source columns")


        client_name_col = find_matching_column(
            df.columns,
            [
                "Client Name", "Policyholder", "Policy Holder", "Investor Name",
                "Insured Name", "Life Assured", "Name"
            ]
        )
        contract_col = find_matching_column(
            df.columns,
            [
                "Contract Number", "Policy Number", "Policy No", "Plan Number",
                "Account Number", "Account No", "Membership Number"
            ]
        )
        total_commission_col = find_matching_column(
            df.columns,
            [
                "Total Commission", "Commission", "Commission Amount",
                "Net Commission", "Amount", "Amount Due"
            ]
        )
        planner_col = find_matching_column(
            df.columns,
            [
                "Planner", "Broker", "Broker Name", "Adviser", "Advisor",
                "Intermediary", "Agent"
            ]
        )


        print(f"   client_name_col       : {client_name_col}")
        print(f"   contract_col          : {contract_col}")
        print(f"   total_commission_col  : {total_commission_col}")
        print(f"   planner_col           : {planner_col}")


        if contract_col is None or total_commission_col is None:
            print("🔴 Required Old Mutual Life Invest columns could not be identified.")
            return None


        # -------------------------------------------------------------------
        # Step 4: Build normalized output df
        # -------------------------------------------------------------------
        print("🟦 Step 4: Build normalized output df")
        normalized_df = pd.DataFrame()


        if client_name_col:
            normalized_df["client_name"] = (
                df[client_name_col]
                .apply(normalize_text_preserve_case)
                .replace("", "tbc")
            )
        else:
            normalized_df["client_name"] = "tbc"


        normalized_df["product_house"] = "old_mutual_life_invest"
        normalized_df["commission_month"] = comm_month


        contract_series = df[contract_col]
        normalized_df["contract_number"] = contract_series.apply(
            lambda x: ""
            if pd.isna(x)
            else (
                str(int(x))
                if isinstance(x, (int, float)) and float(x).is_integer()
                else normalize_text_preserve_case(x)
            )
        )


        normalized_df["total_commission"] = coerce_series_to_numeric(df[total_commission_col])


        if planner_col:
            normalized_df["planner"] = (
                df[planner_col]
                .apply(normalize_text_preserve_case)
                .replace("", "tbc")
            )
        else:
            normalized_df["planner"] = "tbc"


        print(f"   normalized_df shape before filters: {normalized_df.shape}")


        # -------------------------------------------------------------------
        # Step 5: Filter invalid rows
        # -------------------------------------------------------------------
        print("🟦 Step 5: Filter invalid rows")
        normalized_df = normalized_df[
            normalized_df["contract_number"].apply(lambda x: normalize_text(x) != "")
        ].copy()


        normalized_df = normalized_df[
            normalized_df["total_commission"].notna()
        ].copy()


        normalized_df = normalized_df[
            ~normalized_df["contract_number"].apply(
                lambda x: normalize_text(x) in ["total", "subtotal", "sub total", "totals:"]
            )
        ].copy()


        normalized_df = normalized_df[
            ~normalized_df["client_name"].apply(
                lambda x: normalize_text(x) in ["total", "subtotal", "sub total", "totals:"]
            )
        ].copy()


        normalized_df = normalized_df.reset_index(drop=True)
        print(f"   normalized_df shape after filters: {normalized_df.shape}")


        if normalized_df.empty:
            print("🔴 No usable Old Mutual Life Invest rows remain after filtering.")
            return None


        # -------------------------------------------------------------------
        # Step 6: Reorder final columns
        # -------------------------------------------------------------------
        print("🟦 Step 6: Reorder final columns")
        normalized_df = normalized_df[
            [
                "client_name",
                "product_house",
                "commission_month",
                "contract_number",
                "total_commission",
                "planner",
            ]
        ].copy()


        print("✅ Old Mutual Life Invest normalized read complete")
        print(normalized_df.head(10))


        return normalized_df


    except Exception as e:
        print(f"🔴 Error in read_old_mutual_life_invest_df: {e}")
        print(traceback.format_exc())
        return None

# ---------------------------------------------------------------------------
# 5.2.x Reader: Kaelo
# ---------------------------------------------------------------------------
@register_reader("Kaelo")
def read_kaelo_df(file_path, **kwargs):
    """
    Reader for Kaelo commission statement files.


    Note:
    - the uploaded sample uses a multi-row title area
    - the actual column header row is on Excel row 19
    - Premium Commission is the value that matches the processed output
      from the uploaded reference file


    Expected output columns:
    - client_name
    - product_house
    - commission_month
    - contract_number
    - total_commission
    - planner
    """
    try:
        print("------------------------------------------------------------")
        print("🧾 5.2.x Reader: read_kaelo_df")
        print(f"   file_path: {file_path}")


        # -------------------------------------------------------------------
        # Step 0: Get passed variables
        # -------------------------------------------------------------------
        print("🟦 Step 0: Get passed variables")
        comm_month = kwargs.get("comm_month")
        print(f"   comm_month: {comm_month}")


        # -------------------------------------------------------------------
        # Step 1: Read best matching sheet
        # -------------------------------------------------------------------
        print("🟦 Step 1: Read best matching Kaelo sheet")


        required_column_candidates = [
            ["Policy Number"],
            ["Insured"],
            ["Adviser", "Advisor"],
            ["Premium Commission"],
        ]


        source_sheet_name, df = read_best_matching_sheet_by_required_columns(
            file_path=file_path,
            required_column_candidates=required_column_candidates,
            header=18,
            exclude_sheet_names=["processed_data"]
        )


        print(f"   source_sheet_name: {source_sheet_name}")


        if df is None or df.empty:
            print("🔴 No usable Kaelo sheet could be identified.")
            return None


        print(f"   raw shape: {df.shape}")
        print(f"   columns after normalization: {list(df.columns)}")


        # -------------------------------------------------------------------
        # Step 2: Drop blank rows
        # -------------------------------------------------------------------
        print("🟦 Step 2: Drop blank rows")
        df = drop_fully_blank_rows(df)
        print(f"   shape after blank-row cleanup: {df.shape}")


        if df.empty:
            print("🔴 No usable Kaelo rows remain after blank-row cleanup.")
            return None


        # -------------------------------------------------------------------
        # Step 3: Detect source columns
        # -------------------------------------------------------------------
        print("🟦 Step 3: Detect source columns")


        client_name_col = find_matching_column(
            df.columns,
            ["Insured", "Client Name", "Policyholder", "Policy Holder", "Name"]
        )
        contract_col = find_matching_column(
            df.columns,
            ["Policy Number", "Policy No", "Contract Number"]
        )
        total_commission_col = find_matching_column(
            df.columns,
            ["Premium Commission", "Commission", "Net Payable"]
        )
        planner_col = find_matching_column(
            df.columns,
            ["Adviser", "Advisor", "Planner", "Broker", "Agent"]
        )


        print(f"   client_name_col       : {client_name_col}")
        print(f"   contract_col          : {contract_col}")
        print(f"   total_commission_col  : {total_commission_col}")
        print(f"   planner_col           : {planner_col}")


        if contract_col is None or total_commission_col is None:
            print("🔴 Required Kaelo columns could not be identified.")
            return None


        # -------------------------------------------------------------------
        # Step 4: Build normalized output df
        # -------------------------------------------------------------------
        print("🟦 Step 4: Build normalized output df")
        normalized_df = pd.DataFrame()


        if client_name_col:
            normalized_df["client_name"] = (
                df[client_name_col]
                .apply(normalize_text_preserve_case)
                .replace("", "tbc")
            )
        else:
            normalized_df["client_name"] = "tbc"


        normalized_df["product_house"] = "kaelo"
        normalized_df["commission_month"] = comm_month


        normalized_df["contract_number"] = (
            df[contract_col]
            .apply(normalize_text_preserve_case)
            .astype(str)
            .str.strip()
        )


        normalized_df["total_commission"] = coerce_series_to_numeric(df[total_commission_col])


        if planner_col:
            normalized_df["planner"] = (
                df[planner_col]
                .apply(normalize_text_preserve_case)
                .replace("", "tbc")
            )
        else:
            normalized_df["planner"] = "tbc"


        print(f"   normalized_df shape before filters: {normalized_df.shape}")


        # -------------------------------------------------------------------
        # Step 5: Filter invalid rows
        # -------------------------------------------------------------------
        print("🟦 Step 5: Filter invalid rows")
        normalized_df = normalized_df[
            normalized_df["contract_number"].apply(lambda x: normalize_text(x) != "")
        ].copy()


        normalized_df = normalized_df[
            normalized_df["total_commission"].notna()
        ].copy()


        normalized_df = normalized_df[
            ~normalized_df["contract_number"].apply(
                lambda x: normalize_text(x) in ["total", "subtotal", "sub total", "totals:"]
            )
        ].copy()


        normalized_df = normalized_df[
            ~normalized_df["client_name"].apply(
                lambda x: normalize_text(x) in ["total", "subtotal", "sub total", "totals:"]
            )
        ].copy()


        normalized_df = normalized_df.reset_index(drop=True)
        print(f"   normalized_df shape after filters: {normalized_df.shape}")


        if normalized_df.empty:
            print("🔴 No usable Kaelo rows remain after filtering.")
            return None


        # -------------------------------------------------------------------
        # Step 6: Reorder final columns
        # -------------------------------------------------------------------
        print("🟦 Step 6: Reorder final columns")
        normalized_df = normalized_df[
            [
                "client_name",
                "product_house",
                "commission_month",
                "contract_number",
                "total_commission",
                "planner",
            ]
        ].copy()


        print("✅ Kaelo normalized read complete")
        print(normalized_df.head(10))


        return normalized_df


    except Exception as e:
        print(f"🔴 Error in read_kaelo_df: {e}")
        print(traceback.format_exc())
        return None

# ---------------------------------------------------------------------------
# 5.2.x Reader: MUA
# ---------------------------------------------------------------------------
@register_reader("MUA")
def read_mua_df(file_path, **kwargs):
    """
    Reader for MUA commission statement files.


    Practical note:
    - I could not access an MUA sample workbook from the currently mounted
      files in this environment, so this version is written as a broad,
      defensive fixed-header reader
    - it will also work if the workbook has already been converted into a
      processed_data-style layout


    Expected output columns:
    - client_name
    - product_house
    - commission_month
    - contract_number
    - total_commission
    - planner
    """
    try:
        print("------------------------------------------------------------")
        print("🧾 5.2.x Reader: read_mua_df")
        print(f"   file_path: {file_path}")


        # -------------------------------------------------------------------
        # Step 0: Get passed variables
        # -------------------------------------------------------------------
        print("🟦 Step 0: Get passed variables")
        comm_month = kwargs.get("comm_month")
        print(f"   comm_month: {comm_month}")


        # -------------------------------------------------------------------
        # Step 1: Read best matching sheet
        # -------------------------------------------------------------------
        print("🟦 Step 1: Read best matching MUA sheet")


        required_column_candidates = [
            ["Client Name", "Insured", "Policyholder", "Name"],
            ["Contract Number", "Policy Number", "Policy No", "Account Number"],
            ["Total Commission", "Commission", "Commission Amount", "Amount", "Net Payable"],
        ]


        source_sheet_name, df = read_best_matching_sheet_by_required_columns(
            file_path=file_path,
            required_column_candidates=required_column_candidates,
            header=0,
            exclude_sheet_names=[]
        )


        print(f"   source_sheet_name: {source_sheet_name}")


        if df is None or df.empty:
            print("🔴 No usable MUA sheet could be identified.")
            return None


        print(f"   raw shape: {df.shape}")
        print(f"   columns after normalization: {list(df.columns)}")


        # -------------------------------------------------------------------
        # Step 2: Drop blank rows
        # -------------------------------------------------------------------
        print("🟦 Step 2: Drop blank rows")
        df = drop_fully_blank_rows(df)
        print(f"   shape after blank-row cleanup: {df.shape}")


        if df.empty:
            print("🔴 No usable MUA rows remain after blank-row cleanup.")
            return None


        # -------------------------------------------------------------------
        # Step 3: Detect source columns
        # -------------------------------------------------------------------
        print("🟦 Step 3: Detect source columns")


        client_name_col = find_matching_column(
            df.columns,
            [
                "Client Name", "Insured", "Insured Name", "Policyholder",
                "Policy Holder", "Member Name", "Name"
            ]
        )
        contract_col = find_matching_column(
            df.columns,
            [
                "Contract Number", "Policy Number", "Policy No",
                "Account Number", "Membership Number"
            ]
        )
        total_commission_col = find_matching_column(
            df.columns,
            [
                "Total Commission", "Commission", "Commission Amount",
                "Amount", "Net Payable", "Amount Due"
            ]
        )
        planner_col = find_matching_column(
            df.columns,
            [
                "Planner", "Broker", "Broker Name", "Adviser", "Advisor",
                "Agent", "Intermediary"
            ]
        )


        print(f"   client_name_col       : {client_name_col}")
        print(f"   contract_col          : {contract_col}")
        print(f"   total_commission_col  : {total_commission_col}")
        print(f"   planner_col           : {planner_col}")


        if contract_col is None or total_commission_col is None:
            print("🔴 Required MUA columns could not be identified.")
            return None


        # -------------------------------------------------------------------
        # Step 4: Build normalized output df
        # -------------------------------------------------------------------
        print("🟦 Step 4: Build normalized output df")
        normalized_df = pd.DataFrame()


        if client_name_col:
            normalized_df["client_name"] = (
                df[client_name_col]
                .apply(normalize_text_preserve_case)
                .replace("", "tbc")
            )
        else:
            normalized_df["client_name"] = "tbc"


        normalized_df["product_house"] = "mua"
        normalized_df["commission_month"] = comm_month


        contract_series = df[contract_col]
        normalized_df["contract_number"] = contract_series.apply(
            lambda x: ""
            if pd.isna(x)
            else (
                str(int(x))
                if isinstance(x, (int, float)) and float(x).is_integer()
                else normalize_text_preserve_case(x)
            )
        )


        normalized_df["total_commission"] = coerce_series_to_numeric(df[total_commission_col])


        if planner_col:
            normalized_df["planner"] = (
                df[planner_col]
                .apply(normalize_text_preserve_case)
                .replace("", "tbc")
            )
        else:
            normalized_df["planner"] = "tbc"


        print(f"   normalized_df shape before filters: {normalized_df.shape}")


        # -------------------------------------------------------------------
        # Step 5: Filter invalid rows
        # -------------------------------------------------------------------
        print("🟦 Step 5: Filter invalid rows")
        normalized_df = normalized_df[
            normalized_df["contract_number"].apply(lambda x: normalize_text(x) != "")
        ].copy()


        normalized_df = normalized_df[
            normalized_df["total_commission"].notna()
        ].copy()


        normalized_df = normalized_df[
            ~normalized_df["contract_number"].apply(
                lambda x: normalize_text(x) in ["total", "subtotal", "sub total", "totals:"]
            )
        ].copy()


        normalized_df = normalized_df[
            ~normalized_df["client_name"].apply(
                lambda x: normalize_text(x) in ["total", "subtotal", "sub total", "totals:"]
            )
        ].copy()


        normalized_df = normalized_df.reset_index(drop=True)
        print(f"   normalized_df shape after filters: {normalized_df.shape}")


        if normalized_df.empty:
            print("🔴 No usable MUA rows remain after filtering.")
            return None


        # -------------------------------------------------------------------
        # Step 6: Reorder final columns
        # -------------------------------------------------------------------
        print("🟦 Step 6: Reorder final columns")
        normalized_df = normalized_df[
            [
                "client_name",
                "product_house",
                "commission_month",
                "contract_number",
                "total_commission",
                "planner",
            ]
        ].copy()


        print("✅ MUA normalized read complete")
        print(normalized_df.head(10))


        return normalized_df


    except Exception as e:
        print(f"🔴 Error in read_mua_df: {e}")
        print(traceback.format_exc())
        return None

# ---------------------------------------------------------------------------
# 5.3 Range-Based / Openpyxl Reader Functions
# ---------------------------------------------------------------------------

# ---------------------------------------------------------------------------
# 5.4 Special-Layout Reader Functions
# ---------------------------------------------------------------------------
# ---------------------------------------------------------------------------
# 5.4.1 Reader: Momentum_Mandy_PDF
# ---------------------------------------------------------------------------
@register_reader("Momentum_Mandy_PDF")
def read_momentum_mandy_pdf_df(file_path, **kwargs):
    """
    Reader for Momentum Mandy PDF commission statements.

    Expected PDF structure:
    - first page contains statement metadata
    - each data row is rendered vertically as:
        contract number
        client
        broker
        broker code
        commission
    - section headings define the current product / payment context
    - some sections continue across pages without repeating headings

    Returns normalized columns:
    - client_name
    - product_house
    - commission_month
    - contract_number
    - total_commission
    - planner
    """
    try:
        print("------------------------------------------------------------")
        print("🧾 5.4.1 Reader: read_momentum_mandy_pdf_df")
        print(f"   file_path: {file_path}")

        # -------------------------------------------------------------------
        # Step 0: Get passed variables
        # -------------------------------------------------------------------
        print("🟦 Step 0: Get passed variables")
        comm_month = kwargs.get("comm_month")
        comm_tables_main_df = kwargs.get("comm_tables_main_df")

        print(f"   comm_month: {comm_month}")
        print(
            f"   comm_tables_main_df shape: "
            f"{comm_tables_main_df.shape if isinstance(comm_tables_main_df, pd.DataFrame) else None}"
        )

        # -------------------------------------------------------------------
        # Step 1: Read raw PDF text lines
        # -------------------------------------------------------------------
        print("🟦 Step 1: Read raw PDF text lines")
        pdf_lines = read_raw_pdf_text_lines(file_path)
        print(f"   extracted line count: {len(pdf_lines)}")

        if not pdf_lines:
            print("🔴 No PDF text lines were extracted.")
            return None

        # -------------------------------------------------------------------
        # Step 2: Define heading and pattern logic
        # -------------------------------------------------------------------
        print("🟦 Step 2: Define heading and pattern logic")

        product_house_map = {
            "Investo": "investo",
            "Momentum Health": "momentum_health",
            "Momentum Short-Term": "momentum_short_term",
            "Momentum Wealth": "momentum_wealth",
            "Myriad": "myriad",
        }

        section_level_1_values = {
            "As And When",
            "Ongoing",
            "Upfront",
        }

        section_level_2_values = {
            "Premium received",
            "Trail Commission",
            "Alteration",
            "Payment",
            "Unpaid",
            "Advisory Fee",
            "Renewal",
            "CPI",
        }

        header_lines_to_skip = {
            "Contract number",
            "Client",
            "Broker",
            "Broker Code",
            "Commission & Fees",
        }

        contract_pattern = re.compile(r"^(?:(?:[A-Z]{1,3}|\d{1,3})\s+)?\d{5,12}$")
        broker_code_pattern = re.compile(r"^\d{6}$")
        amount_pattern = re.compile(r"^-?R[\d,]+\.\d{2}$")
        subtotal_pattern = re.compile(r"^(.+?)\s+SubTotal$", re.IGNORECASE)

        # -------------------------------------------------------------------
        # Step 3: Parse metadata and transactional rows
        # -------------------------------------------------------------------
        print("🟦 Step 3: Parse PDF line structure")

        broker_from_header = None
        broker_code_from_header = None

        parsed_rows = []
        parsed_subtotals = {}
        parsed_month_total = None
        parsed_total = None

        current_product_house = None
        current_section_level_1 = None
        current_section_level_2 = None

        i = 0
        while i < len(pdf_lines):
            line = str(pdf_lines[i]).strip()

            if line == "":
                i += 1
                continue

            # ---------------------------------------------------------------
            # Step 3.1: Header metadata
            # ---------------------------------------------------------------
            if line == "Broker:" and i + 1 < len(pdf_lines):
                broker_from_header = str(pdf_lines[i + 1]).strip()
                print(f"   broker_from_header: {broker_from_header}")
                i += 2
                continue

            if line == "Broker code:" and i + 1 < len(pdf_lines):
                broker_code_from_header = str(pdf_lines[i + 1]).strip()
                print(f"   broker_code_from_header: {broker_code_from_header}")
                i += 2
                continue

            # ---------------------------------------------------------------
            # Step 3.2: Footer / noise lines
            # ---------------------------------------------------------------
            if re.fullmatch(r"\d{1,2}/\d{1,2}/\d{2},.*", line):
                i += 1
                continue

            if re.fullmatch(r"\d+/\d+", line):
                i += 1
                continue

            if line.startswith("http://") or line.startswith("https://"):
                i += 1
                continue

            if line in {
                "Production statement",
                "Production statement:",
                "Financial month:",
                "Broker house:",
                "Broker house code:",
            }:
                i += 1
                continue

            if line.startswith("Broker house code:"):
                i += 1
                continue

            if line.startswith("January") and "broker house code" in line:
                i += 1
                continue

            if line in {"January 2026", "HOWARD JOHN CHARLES SILK"}:
                i += 1
                continue

            # ---------------------------------------------------------------
            # Step 3.3: Structural headings
            # ---------------------------------------------------------------
            if line in product_house_map:
                current_product_house = product_house_map[line]
                print(f"   current_product_house set to: {current_product_house}")
                i += 1
                continue

            if line in section_level_1_values:
                current_section_level_1 = line
                i += 1
                continue

            if line in section_level_2_values:
                current_section_level_2 = line
                i += 1
                continue

            if line in header_lines_to_skip:
                i += 1
                continue

            # ---------------------------------------------------------------
            # Step 3.4: Section subtotals
            # ---------------------------------------------------------------
            subtotal_match = subtotal_pattern.match(line)
            if subtotal_match:
                subtotal_label = subtotal_match.group(1).strip()
                next_line = str(pdf_lines[i + 1]).strip() if i + 1 < len(pdf_lines) else ""

                if amount_pattern.match(next_line):
                    subtotal_value = float(next_line.replace("R", "").replace(",", ""))
                    parsed_subtotals[subtotal_label] = subtotal_value
                    print(f"   parsed subtotal -> {subtotal_label}: {subtotal_value}")
                    i += 2
                    continue

                i += 1
                continue

            # ---------------------------------------------------------------
            # Step 3.5: Statement totals
            # ---------------------------------------------------------------
            if line == "Month total":
                next_line = str(pdf_lines[i + 1]).strip() if i + 1 < len(pdf_lines) else ""
                if amount_pattern.match(next_line):
                    parsed_month_total = float(next_line.replace("R", "").replace(",", ""))
                    print(f"   parsed_month_total: {parsed_month_total}")
                    i += 2
                    continue

            if line == "Total":
                next_line = str(pdf_lines[i + 1]).strip() if i + 1 < len(pdf_lines) else ""
                if amount_pattern.match(next_line):
                    parsed_total = float(next_line.replace("R", "").replace(",", ""))
                    print(f"   parsed_total: {parsed_total}")
                    i += 2
                    continue

            # ---------------------------------------------------------------
            # Step 3.6: Transaction rows
            # ---------------------------------------------------------------
            if contract_pattern.match(line) and i + 4 < len(pdf_lines):
                contract_number = line
                client_name = str(pdf_lines[i + 1]).strip()
                planner = str(pdf_lines[i + 2]).strip()
                broker_code = str(pdf_lines[i + 3]).strip()
                amount_text = str(pdf_lines[i + 4]).strip()

                if broker_code_pattern.match(broker_code) and amount_pattern.match(amount_text):
                    total_commission = float(amount_text.replace("R", "").replace(",", ""))

                    parsed_rows.append({
                        "source_product_house": current_product_house,
                        "source_section_level_1": current_section_level_1,
                        "source_section_level_2": current_section_level_2,
                        "contract_number": contract_number,
                        "client_name": client_name,
                        "planner": planner if planner != "" else broker_from_header,
                        "broker_code": broker_code if broker_code != "" else broker_code_from_header,
                        "total_commission": total_commission,
                    })

                    i += 5
                    continue

            i += 1

        print(f"   parsed transactional rows: {len(parsed_rows)}")

        if len(parsed_rows) == 0:
            print("🔴 No transaction rows could be parsed from the Momentum PDF.")
            return None

        # -------------------------------------------------------------------
        # Step 4: Convert parsed rows to DataFrame
        # -------------------------------------------------------------------
        print("🟦 Step 4: Convert parsed rows to DataFrame")
        parsed_df = pd.DataFrame(parsed_rows)
        print(f"   parsed_df shape: {parsed_df.shape}")
        print(parsed_df.head(10))

        # -------------------------------------------------------------------
        # Step 5: Validate parsed totals
        # -------------------------------------------------------------------
        print("🟦 Step 5: Validate parsed totals against PDF subtotals")

        actual_subtotals = (
            parsed_df.groupby("source_product_house", dropna=False)["total_commission"]
            .sum()
            .round(2)
            .to_dict()
        )

        reverse_product_house_map = {v: k for k, v in product_house_map.items()}

        for source_key, actual_total in actual_subtotals.items():
            display_label = reverse_product_house_map.get(source_key)
            expected_total = parsed_subtotals.get(display_label)

            print(
                f"   subtotal check | source_key={source_key} | "
                f"display_label={display_label} | actual_total={actual_total} | "
                f"expected_total={expected_total}"
            )

        parsed_grand_total = round(parsed_df["total_commission"].sum(), 2)
        print(f"   parsed_grand_total: {parsed_grand_total}")
        print(f"   PDF month total   : {parsed_month_total}")
        print(f"   PDF total         : {parsed_total}")

        # -------------------------------------------------------------------
        # Step 6: Build normalized output DataFrame
        # -------------------------------------------------------------------
        print("🟦 Step 6: Build normalized output DataFrame")

        normalized_df = pd.DataFrame()
        normalized_df["client_name"] = parsed_df["client_name"].astype(str).str.strip()
        normalized_df["product_house"] = "momentum_mandy"
        normalized_df["commission_month"] = comm_month
        normalized_df["contract_number"] = parsed_df["contract_number"].astype(str).str.strip()
        normalized_df["total_commission"] = pd.to_numeric(parsed_df["total_commission"], errors="coerce")
        normalized_df["planner"] = parsed_df["planner"].fillna("").astype(str).str.strip().replace("", broker_from_header if broker_from_header else "tbc")

        print(f"   normalized_df shape before filters: {normalized_df.shape}")
        print(normalized_df.head(10))

        # -------------------------------------------------------------------
        # Step 7: Filter invalid rows
        # -------------------------------------------------------------------
        print("🟦 Step 7: Filter invalid rows")

        normalized_df = normalized_df[
            normalized_df["contract_number"].astype(str).str.strip() != ""
        ].copy()

        normalized_df = normalized_df[
            normalized_df["total_commission"].notna()
        ].copy()

        normalized_df = normalized_df.reset_index(drop=True)
        print(f"   normalized_df shape after filters: {normalized_df.shape}")

        if normalized_df.empty:
            print("🔴 No usable Momentum Mandy PDF rows remain after filtering.")
            return None

        # -------------------------------------------------------------------
        # Step 8: Reorder final columns
        # -------------------------------------------------------------------
        print("🟦 Step 8: Reorder final columns")

        normalized_df = normalized_df[
            [
                "client_name",
                "product_house",
                "commission_month",
                "contract_number",
                "total_commission",
                "planner",
            ]
        ].copy()

        print("✅ Momentum Mandy PDF normalized read complete")
        print(normalized_df.head(10))

        return normalized_df

    except Exception as e:
        print(f"🔴 Error in read_momentum_mandy_pdf_df: {e}")
        print(traceback.format_exc())
        return None


# ---------------------------------------------------------------------------
# Reader Functions Patch
# ---------------------------------------------------------------------------
# Purpose:
# Paste these reader functions into process_engine.py under the
# existing "5. READER FUNCTIONS" area.
#
# Notes:
# - These readers are written to use the existing utilities already present in
#   process_engine.py, in line with the current project structure.
# - No new helper utilities are introduced.
# - Sanlam and Sirago are written as fallback readers against an already-
#   normalized sheet because the uploaded source examples only contained
#   processed_data-style output, not the original raw statement layouts.
# ---------------------------------------------------------------------------

# ---------------------------------------------------------------------------
# 5.4.2 Reader: Zestlife
# ---------------------------------------------------------------------------
@register_reader("Zestlife")
def read_zestlife_df(file_path, **kwargs):
    """
    Reader for Zestlife commission statement files.

    Expected source layout:
    - workbook contains a rendered remittance-style sheet
    - usable header row sits on Excel row 18
    - data rows begin below that header

    Returns normalized columns:
    - client_name
    - product_house
    - commission_month
    - contract_number
    - total_commission
    - planner
    """
    try:
        print("------------------------------------------------------------")
        print("🧾 5.4.2 Reader: read_zestlife_df")
        print(f"   file_path: {file_path}")

        # -------------------------------------------------------------------
        # Step 0: Get passed variables
        # -------------------------------------------------------------------
        print("🟦 Step 0: Get passed variables")
        comm_month = kwargs.get("comm_month")
        comm_tables_main_df = kwargs.get("comm_tables_main_df")

        print(f"   comm_month: {comm_month}")
        print(
            f"   comm_tables_main_df shape: "
            f"{comm_tables_main_df.shape if isinstance(comm_tables_main_df, pd.DataFrame) else None}"
        )

        # -------------------------------------------------------------------
        # Step 1: Read best matching raw sheet
        # -------------------------------------------------------------------
        print("🟦 Step 1: Read best matching source sheet")
        source_sheet_name, df = read_best_matching_sheet_by_required_columns(
            file_path=file_path,
            required_column_candidates=[
                ["Policy Number"],
                ["Full Names", "Full Name"],
                ["Surname"],
                ["Sales Agent"],
                ["Net Payable"],
            ],
            header=17,
            exclude_sheet_names=["processed_data"],
        )

        print(f"   source_sheet_name: {source_sheet_name}")
        print(f"   raw shape: {df.shape if isinstance(df, pd.DataFrame) else None}")

        if df is None or df.empty:
            print("🔴 No usable Zestlife sheet could be identified.")
            return None

        # -------------------------------------------------------------------
        # Step 2: Detect source columns
        # -------------------------------------------------------------------
        print("🟦 Step 2: Detect source columns")
        contract_col = find_matching_column(df.columns, ["Policy Number"])
        first_name_col = find_matching_column(df.columns, ["Full Names", "Full Name"])
        surname_col = find_matching_column(df.columns, ["Surname"])
        planner_col = find_matching_column(df.columns, ["Sales Agent"])
        total_commission_col = find_matching_column(df.columns, ["Net Payable"])
        status_col = find_matching_column(df.columns, ["Status"])
        show_col = find_matching_column(df.columns, ["unnamed_1", "show"], allow_contains=False)

        print(f"   contract_col          : {contract_col}")
        print(f"   first_name_col        : {first_name_col}")
        print(f"   surname_col           : {surname_col}")
        print(f"   planner_col           : {planner_col}")
        print(f"   total_commission_col  : {total_commission_col}")
        print(f"   status_col            : {status_col}")
        print(f"   show_col              : {show_col}")

        if contract_col is None or total_commission_col is None:
            print("🔴 Required Zestlife columns could not be identified.")
            return None

        # -------------------------------------------------------------------
        # Step 3: Build normalized output df
        # -------------------------------------------------------------------
        print("🟦 Step 3: Build normalized output df")
        normalized_df = pd.DataFrame()

        if first_name_col and surname_col:
            normalized_df["client_name"] = (
                df[first_name_col].apply(normalize_text_preserve_case)
                + " "
                + df[surname_col].apply(normalize_text_preserve_case)
            ).str.replace(r"\s+", " ", regex=True).str.strip()
        elif first_name_col:
            normalized_df["client_name"] = df[first_name_col].apply(normalize_text_preserve_case)
        else:
            normalized_df["client_name"] = ""

        normalized_df["product_house"] = "zestlife"
        normalized_df["commission_month"] = comm_month
        normalized_df["contract_number"] = df[contract_col].apply(normalize_text_preserve_case)
        normalized_df["total_commission"] = coerce_series_to_numeric(df[total_commission_col])

        if planner_col:
            normalized_df["planner"] = (
                df[planner_col]
                .apply(normalize_text_preserve_case)
                .replace("", "UNIFY (PTY) LTD")
            )
        else:
            normalized_df["planner"] = "UNIFY (PTY) LTD"

        print(f"   normalized_df shape before filters: {normalized_df.shape}")

        # -------------------------------------------------------------------
        # Step 4: Filter invalid rows
        # -------------------------------------------------------------------
        print("🟦 Step 4: Filter invalid rows")

        if show_col is not None:
            show_mask = df[show_col].apply(lambda x: normalize_text(x) in ["show", ""])
            normalized_df = normalized_df[show_mask].copy()
            print(f"   shape after show filter: {normalized_df.shape}")

        if status_col is not None:
            status_mask = df[status_col].apply(lambda x: normalize_text(x) not in ["", "policy details"])
            normalized_df = normalized_df[status_mask].copy()
            print(f"   shape after status filter: {normalized_df.shape}")

        normalized_df = normalized_df[
            normalized_df["contract_number"].apply(lambda x: normalize_text(x) != "")
        ].copy()

        normalized_df = normalized_df[
            normalized_df["total_commission"].notna()
        ].copy()

        normalized_df = normalized_df.reset_index(drop=True)
        print(f"   normalized_df shape after filters: {normalized_df.shape}")

        if normalized_df.empty:
            print("🔴 No usable Zestlife rows remain after filtering.")
            return None

        # -------------------------------------------------------------------
        # Step 5: Reorder final columns
        # -------------------------------------------------------------------
        print("🟦 Step 5: Reorder final columns")
        normalized_df = normalized_df[
            [
                "client_name",
                "product_house",
                "commission_month",
                "contract_number",
                "total_commission",
                "planner",
            ]
        ].copy()

        print("✅ Zestlife normalized read complete")
        print(normalized_df.head(10))

        return normalized_df

    except Exception as e:
        print(f"🔴 Error in read_zestlife_df: {e}")
        print(traceback.format_exc())
        return None


# ---------------------------------------------------------------------------
# 5.4.3 Reader: SAU
# ---------------------------------------------------------------------------
@register_reader("SAU")
def read_sau_df(file_path, **kwargs):
    """
    Reader for SAU commission statement files.

    Expected source layout:
    - workbook contains multiple summary tabs plus a Detail tab
    - usable header row sits on Excel row 11
    - Detail sheet contains line-item data

    Current logic maps total_commission from 'Total Due' because that aligns
    with the net payable commission value. If you later decide you want gross
    including VAT instead, switch the source column to 'Total Incl. VAT'.
    """
    try:
        print("------------------------------------------------------------")
        print("🧾 5.4.3 Reader: read_sau_df")
        print(f"   file_path: {file_path}")

        # -------------------------------------------------------------------
        # Step 0: Get passed variables
        # -------------------------------------------------------------------
        print("🟦 Step 0: Get passed variables")
        comm_month = kwargs.get("comm_month")
        comm_tables_main_df = kwargs.get("comm_tables_main_df")

        print(f"   comm_month: {comm_month}")
        print(
            f"   comm_tables_main_df shape: "
            f"{comm_tables_main_df.shape if isinstance(comm_tables_main_df, pd.DataFrame) else None}"
        )

        # -------------------------------------------------------------------
        # Step 1: Read best matching source sheet
        # -------------------------------------------------------------------
        print("🟦 Step 1: Read best matching source sheet")
        source_sheet_name, df = read_best_matching_sheet_by_required_columns(
            file_path=file_path,
            required_column_candidates=[
                ["Client Name"],
                ["Policy Number"],
                ["Total Due"],
                ["Agent"],
            ],
            header=10,
            exclude_sheet_names=["processed_data"],
        )

        print(f"   source_sheet_name: {source_sheet_name}")
        print(f"   raw shape: {df.shape if isinstance(df, pd.DataFrame) else None}")

        if df is None or df.empty:
            print("🔴 No usable SAU sheet could be identified.")
            return None

        # -------------------------------------------------------------------
        # Step 2: Detect source columns
        # -------------------------------------------------------------------
        print("🟦 Step 2: Detect source columns")
        client_name_col = find_matching_column(df.columns, ["Client Name"])
        contract_col = find_matching_column(df.columns, ["Policy Number"])
        total_commission_col = find_matching_column(df.columns, ["Total Due"])
        planner_col = find_matching_column(df.columns, ["Agent"])
        description_col = find_matching_column(df.columns, ["Description"])

        print(f"   client_name_col       : {client_name_col}")
        print(f"   contract_col          : {contract_col}")
        print(f"   total_commission_col  : {total_commission_col}")
        print(f"   planner_col           : {planner_col}")
        print(f"   description_col       : {description_col}")

        if contract_col is None or total_commission_col is None:
            print("🔴 Required SAU columns could not be identified.")
            return None

        # -------------------------------------------------------------------
        # Step 3: Build normalized output df
        # -------------------------------------------------------------------
        print("🟦 Step 3: Build normalized output df")
        normalized_df = pd.DataFrame()

        if client_name_col:
            normalized_df["client_name"] = df[client_name_col].apply(normalize_text_preserve_case)
        else:
            normalized_df["client_name"] = ""

        normalized_df["product_house"] = "sau"
        normalized_df["commission_month"] = comm_month
        normalized_df["contract_number"] = df[contract_col].apply(normalize_text_preserve_case)
        normalized_df["total_commission"] = coerce_series_to_numeric(df[total_commission_col])

        if planner_col:
            normalized_df["planner"] = (
                df[planner_col]
                .apply(normalize_text_preserve_case)
                .replace("", "UNIFY (PTY) LTD")
            )
        else:
            normalized_df["planner"] = "UNIFY (PTY) LTD"

        print(f"   normalized_df shape before filters: {normalized_df.shape}")

        # -------------------------------------------------------------------
        # Step 4: Filter invalid rows
        # -------------------------------------------------------------------
        print("🟦 Step 4: Filter invalid rows")
        if description_col is not None:
            detail_mask = df[description_col].apply(lambda x: normalize_text(x) not in ["", "totals"])
            normalized_df = normalized_df[detail_mask].copy()
            print(f"   shape after description filter: {normalized_df.shape}")

        normalized_df = normalized_df[
            normalized_df["contract_number"].apply(lambda x: normalize_text(x) != "")
        ].copy()

        normalized_df = normalized_df[
            normalized_df["total_commission"].notna()
        ].copy()

        normalized_df = normalized_df.reset_index(drop=True)
        print(f"   normalized_df shape after filters: {normalized_df.shape}")

        if normalized_df.empty:
            print("🔴 No usable SAU rows remain after filtering.")
            return None

        # -------------------------------------------------------------------
        # Step 5: Reorder final columns
        # -------------------------------------------------------------------
        print("🟦 Step 5: Reorder final columns")
        normalized_df = normalized_df[
            [
                "client_name",
                "product_house",
                "commission_month",
                "contract_number",
                "total_commission",
                "planner",
            ]
        ].copy()

        print("✅ SAU normalized read complete")
        print(normalized_df.head(10))

        return normalized_df

    except Exception as e:
        print(f"🔴 Error in read_sau_df: {e}")
        print(traceback.format_exc())
        return None


# ---------------------------------------------------------------------------
# 5.4.4 Reader: Stratum
# ---------------------------------------------------------------------------
@register_reader("Stratum")
def read_stratum_df(file_path, **kwargs):
    """
    Reader for Stratum commission statement files.

    Expected source layout:
    - rendered statement workbook
    - usable header row sits on Excel row 12
    - detailed records follow directly below
    """
    try:
        print("------------------------------------------------------------")
        print("🧾 5.4.4 Reader: read_stratum_df")
        print(f"   file_path: {file_path}")

        # -------------------------------------------------------------------
        # Step 0: Get passed variables
        # -------------------------------------------------------------------
        print("🟦 Step 0: Get passed variables")
        comm_month = kwargs.get("comm_month")
        comm_tables_main_df = kwargs.get("comm_tables_main_df")

        print(f"   comm_month: {comm_month}")
        print(
            f"   comm_tables_main_df shape: "
            f"{comm_tables_main_df.shape if isinstance(comm_tables_main_df, pd.DataFrame) else None}"
        )

        # -------------------------------------------------------------------
        # Step 1: Read best matching source sheet
        # -------------------------------------------------------------------
        print("🟦 Step 1: Read best matching source sheet")
        source_sheet_name, df = read_best_matching_sheet_by_required_columns(
            file_path=file_path,
            required_column_candidates=[
                ["Policy Number"],
                ["Name"],
                ["Surname"],
                ["Financial Advisor"],
                ["Commission", "Total Due"],
            ],
            header=11,
            exclude_sheet_names=["processed_data"],
        )

        print(f"   source_sheet_name: {source_sheet_name}")
        print(f"   raw shape: {df.shape if isinstance(df, pd.DataFrame) else None}")

        if df is None or df.empty:
            print("🔴 No usable Stratum sheet could be identified.")
            return None

        # -------------------------------------------------------------------
        # Step 2: Detect source columns
        # -------------------------------------------------------------------
        print("🟦 Step 2: Detect source columns")
        contract_col = find_matching_column(df.columns, ["Policy Number"])
        first_name_col = find_matching_column(df.columns, ["Name"])
        surname_col = find_matching_column(df.columns, ["Surname"])
        planner_col = find_matching_column(df.columns, ["Financial Advisor"])
        total_commission_col = find_matching_column(df.columns, ["Commission", "Total Due"])
        status_col = find_matching_column(df.columns, ["Status"])

        print(f"   contract_col          : {contract_col}")
        print(f"   first_name_col        : {first_name_col}")
        print(f"   surname_col           : {surname_col}")
        print(f"   planner_col           : {planner_col}")
        print(f"   total_commission_col  : {total_commission_col}")
        print(f"   status_col            : {status_col}")

        if contract_col is None or total_commission_col is None:
            print("🔴 Required Stratum columns could not be identified.")
            return None

        # -------------------------------------------------------------------
        # Step 3: Build normalized output df
        # -------------------------------------------------------------------
        print("🟦 Step 3: Build normalized output df")
        normalized_df = pd.DataFrame()

        if first_name_col and surname_col:
            normalized_df["client_name"] = (
                df[first_name_col].apply(normalize_text_preserve_case)
                + " "
                + df[surname_col].apply(normalize_text_preserve_case)
            ).str.replace(r"\s+", " ", regex=True).str.strip()
        elif first_name_col:
            normalized_df["client_name"] = df[first_name_col].apply(normalize_text_preserve_case)
        else:
            normalized_df["client_name"] = ""

        normalized_df["product_house"] = "stratum"
        normalized_df["commission_month"] = comm_month
        normalized_df["contract_number"] = df[contract_col].apply(normalize_text_preserve_case)
        normalized_df["total_commission"] = coerce_series_to_numeric(df[total_commission_col])

        if planner_col:
            normalized_df["planner"] = (
                df[planner_col]
                .apply(normalize_text_preserve_case)
                .replace("", "UNIFY (PTY) LTD")
            )
        else:
            normalized_df["planner"] = "UNIFY (PTY) LTD"

        print(f"   normalized_df shape before filters: {normalized_df.shape}")

        # -------------------------------------------------------------------
        # Step 4: Filter invalid rows
        # -------------------------------------------------------------------
        print("🟦 Step 4: Filter invalid rows")
        if status_col is not None:
            status_mask = df[status_col].apply(
                lambda x: normalize_text(x) in ["live", "cancelled", "pending cancellation", ""]
            )
            normalized_df = normalized_df[status_mask].copy()
            print(f"   shape after status filter: {normalized_df.shape}")

        normalized_df = normalized_df[
            normalized_df["contract_number"].apply(lambda x: normalize_text(x) != "")
        ].copy()

        normalized_df = normalized_df[
            normalized_df["total_commission"].notna()
        ].copy()

        normalized_df = normalized_df.reset_index(drop=True)
        print(f"   normalized_df shape after filters: {normalized_df.shape}")

        if normalized_df.empty:
            print("🔴 No usable Stratum rows remain after filtering.")
            return None

        # -------------------------------------------------------------------
        # Step 5: Reorder final columns
        # -------------------------------------------------------------------
        print("🟦 Step 5: Reorder final columns")
        normalized_df = normalized_df[
            [
                "client_name",
                "product_house",
                "commission_month",
                "contract_number",
                "total_commission",
                "planner",
            ]
        ].copy()

        print("✅ Stratum normalized read complete")
        print(normalized_df.head(10))

        return normalized_df

    except Exception as e:
        print(f"🔴 Error in read_stratum_df: {e}")
        print(traceback.format_exc())
        return None


# ---------------------------------------------------------------------------
# 5.4.5 Reader: Turnberry
# ---------------------------------------------------------------------------
@register_reader("Turnberry")
def read_turnberry_df(file_path, **kwargs):
    """
    Reader for Turnberry commission statement files.

    Expected source layout:
    - first sheet contains a statement heading block
    - usable header row sits on Excel row 4
    - planner name sits on Excel row 2, column A
    """
    try:
        print("------------------------------------------------------------")
        print("🧾 5.4.5 Reader: read_turnberry_df")
        print(f"   file_path: {file_path}")

        # -------------------------------------------------------------------
        # Step 0: Get passed variables
        # -------------------------------------------------------------------
        print("🟦 Step 0: Get passed variables")
        comm_month = kwargs.get("comm_month")
        comm_tables_main_df = kwargs.get("comm_tables_main_df")

        print(f"   comm_month: {comm_month}")
        print(
            f"   comm_tables_main_df shape: "
            f"{comm_tables_main_df.shape if isinstance(comm_tables_main_df, pd.DataFrame) else None}"
        )

        # -------------------------------------------------------------------
        # Step 1: Read raw first sheet for planner capture
        # -------------------------------------------------------------------
        print("🟦 Step 1: Read raw first sheet for planner capture")
        raw_df = read_raw_first_sheet(file_path, header=None)
        print(f"   raw_df shape: {raw_df.shape}")

        if raw_df is None or raw_df.empty:
            print("🔴 Raw Turnberry sheet returned no data.")
            return None

        planner_value = "UNIFY (PTY) LTD"
        if raw_df.shape[0] > 1 and raw_df.shape[1] > 0:
            planner_candidate = normalize_text_preserve_case(raw_df.iloc[1, 0])
            if planner_candidate != "":
                planner_value = planner_candidate

        print(f"   planner_value: {planner_value}")

        # -------------------------------------------------------------------
        # Step 2: Read best matching structured sheet
        # -------------------------------------------------------------------
        print("🟦 Step 2: Read best matching structured sheet")
        source_sheet_name, df = read_best_matching_sheet_by_required_columns(
            file_path=file_path,
            required_column_candidates=[
                ["Policy No", "Policy Number"],
                ["Commission incl Vat", "Commission"],
                ["Name"],
                ["Surname"],
            ],
            header=3,
            exclude_sheet_names=["processed_data"],
        )

        print(f"   source_sheet_name: {source_sheet_name}")
        print(f"   structured df shape: {df.shape if isinstance(df, pd.DataFrame) else None}")

        if df is None or df.empty:
            print("🔴 No usable Turnberry sheet could be identified.")
            return None

        # -------------------------------------------------------------------
        # Step 3: Detect source columns
        # -------------------------------------------------------------------
        print("🟦 Step 3: Detect source columns")
        contract_col = find_matching_column(df.columns, ["Policy No", "Policy Number"])
        total_commission_col = find_matching_column(df.columns, ["Commission incl Vat", "Commission"])
        full_name_col = find_matching_column(df.columns, ["unnamed_15"], allow_contains=False)
        first_name_col = find_matching_column(df.columns, ["Name"])
        surname_col = find_matching_column(df.columns, ["Surname"])

        print(f"   contract_col          : {contract_col}")
        print(f"   total_commission_col  : {total_commission_col}")
        print(f"   full_name_col         : {full_name_col}")
        print(f"   first_name_col        : {first_name_col}")
        print(f"   surname_col           : {surname_col}")

        if contract_col is None or total_commission_col is None:
            print("🔴 Required Turnberry columns could not be identified.")
            return None

        # -------------------------------------------------------------------
        # Step 4: Build normalized output df
        # -------------------------------------------------------------------
        print("🟦 Step 4: Build normalized output df")
        normalized_df = pd.DataFrame()

        if full_name_col is not None:
            normalized_df["client_name"] = df[full_name_col].apply(normalize_text_preserve_case)
        elif first_name_col and surname_col:
            normalized_df["client_name"] = (
                df[first_name_col].apply(normalize_text_preserve_case)
                + " "
                + df[surname_col].apply(normalize_text_preserve_case)
            ).str.replace(r"\s+", " ", regex=True).str.strip()
        else:
            normalized_df["client_name"] = ""

        normalized_df["product_house"] = "turnberry"
        normalized_df["commission_month"] = comm_month
        normalized_df["contract_number"] = df[contract_col].apply(normalize_text_preserve_case)
        normalized_df["total_commission"] = coerce_series_to_numeric(df[total_commission_col])
        normalized_df["planner"] = planner_value

        print(f"   normalized_df shape before filters: {normalized_df.shape}")

        # -------------------------------------------------------------------
        # Step 5: Filter invalid rows
        # -------------------------------------------------------------------
        print("🟦 Step 5: Filter invalid rows")
        normalized_df = normalized_df[
            normalized_df["contract_number"].apply(lambda x: normalize_text(x) != "")
        ].copy()

        normalized_df = normalized_df[
            normalized_df["total_commission"].notna()
        ].copy()

        normalized_df = normalized_df.reset_index(drop=True)
        print(f"   normalized_df shape after filters: {normalized_df.shape}")

        if normalized_df.empty:
            print("🔴 No usable Turnberry rows remain after filtering.")
            return None

        # -------------------------------------------------------------------
        # Step 6: Reorder final columns
        # -------------------------------------------------------------------
        print("🟦 Step 6: Reorder final columns")
        normalized_df = normalized_df[
            [
                "client_name",
                "product_house",
                "commission_month",
                "contract_number",
                "total_commission",
                "planner",
            ]
        ].copy()

        print("✅ Turnberry normalized read complete")
        print(normalized_df.head(10))

        return normalized_df

    except Exception as e:
        print(f"🔴 Error in read_turnberry_df: {e}")
        print(traceback.format_exc())
        return None


# ---------------------------------------------------------------------------
# 5.4.6 Reader: Santam
# ---------------------------------------------------------------------------
@register_reader("Santam")
def read_santam_df(file_path, **kwargs):
    """
    Reader for Santam commission statement files.

    Expected source layout from the uploaded sample:
    - first sheet has no usable header row
    - each row is a raw fixed-position record
    - contract number sits in column E
    - client name sits in column G
    - net commission sits in column M (negative in the sample)

    Planner currently defaults to UNIFY (PTY) LTD because the uploaded raw
    sample did not contain a readable planner-name field.
    """
    try:
        print("------------------------------------------------------------")
        print("🧾 5.4.6 Reader: read_santam_df")
        print(f"   file_path: {file_path}")

        # -------------------------------------------------------------------
        # Step 0: Get passed variables
        # -------------------------------------------------------------------
        print("🟦 Step 0: Get passed variables")
        comm_month = kwargs.get("comm_month")
        comm_tables_main_df = kwargs.get("comm_tables_main_df")

        print(f"   comm_month: {comm_month}")
        print(
            f"   comm_tables_main_df shape: "
            f"{comm_tables_main_df.shape if isinstance(comm_tables_main_df, pd.DataFrame) else None}"
        )

        # -------------------------------------------------------------------
        # Step 1: Read raw first sheet
        # -------------------------------------------------------------------
        print("🟦 Step 1: Read raw first sheet")
        df = read_raw_first_sheet(file_path, header=None)
        print(f"   raw shape: {df.shape}")

        if df is None or df.empty:
            print("🔴 Raw Santam sheet returned no data.")
            return None

        if df.shape[1] < 13:
            print("🔴 Santam raw layout does not contain the expected 13 columns.")
            return None

        # -------------------------------------------------------------------
        # Step 2: Build normalized output df from fixed positions
        # -------------------------------------------------------------------
        print("🟦 Step 2: Build normalized output df from fixed positions")
        normalized_df = pd.DataFrame()
        normalized_df["client_name"] = df.iloc[:, 6].apply(normalize_text_preserve_case)
        normalized_df["product_house"] = "santam"
        normalized_df["commission_month"] = comm_month
        normalized_df["contract_number"] = df.iloc[:, 4].apply(normalize_text_preserve_case)
        normalized_df["total_commission"] = coerce_series_to_numeric(df.iloc[:, 12]).abs()
        normalized_df["planner"] = "UNIFY (PTY) LTD"

        print(f"   normalized_df shape before filters: {normalized_df.shape}")

        # -------------------------------------------------------------------
        # Step 3: Filter invalid rows
        # -------------------------------------------------------------------
        print("🟦 Step 3: Filter invalid rows")
        normalized_df = normalized_df[
            normalized_df["contract_number"].apply(lambda x: normalize_text(x) != "")
        ].copy()

        normalized_df = normalized_df[
            normalized_df["total_commission"].notna()
        ].copy()

        normalized_df = normalized_df.reset_index(drop=True)
        print(f"   normalized_df shape after filters: {normalized_df.shape}")

        if normalized_df.empty:
            print("🔴 No usable Santam rows remain after filtering.")
            return None

        # -------------------------------------------------------------------
        # Step 4: Reorder final columns
        # -------------------------------------------------------------------
        print("🟦 Step 4: Reorder final columns")
        normalized_df = normalized_df[
            [
                "client_name",
                "product_house",
                "commission_month",
                "contract_number",
                "total_commission",
                "planner",
            ]
        ].copy()

        print("✅ Santam normalized read complete")
        print(normalized_df.head(10))

        return normalized_df

    except Exception as e:
        print(f"🔴 Error in read_santam_df: {e}")
        print(traceback.format_exc())
        return None


# ---------------------------------------------------------------------------
# 5.4.7 Reader: Sanlam
# ---------------------------------------------------------------------------
@register_reader("Sanlam")
def read_sanlam_df(file_path, **kwargs):
    """
    Fallback reader for Sanlam files based on the uploaded example.

    Important:
    - the uploaded Sanlam reference workbook only contained an already-
      normalized processed_data-style sheet
    - because of that, this reader currently re-imports a normalized sheet
      rather than parsing a raw Sanlam statement layout
    """
    try:
        print("------------------------------------------------------------")
        print("🧾 5.4.7 Reader: read_sanlam_df")
        print(f"   file_path: {file_path}")

        # -------------------------------------------------------------------
        # Step 0: Get passed variables
        # -------------------------------------------------------------------
        print("🟦 Step 0: Get passed variables")
        comm_month = kwargs.get("comm_month")
        comm_tables_main_df = kwargs.get("comm_tables_main_df")

        print(f"   comm_month: {comm_month}")
        print(
            f"   comm_tables_main_df shape: "
            f"{comm_tables_main_df.shape if isinstance(comm_tables_main_df, pd.DataFrame) else None}"
        )

        # -------------------------------------------------------------------
        # Step 1: Read best matching normalized sheet
        # -------------------------------------------------------------------
        print("🟦 Step 1: Read best matching normalized sheet")
        source_sheet_name, df = read_best_matching_sheet_by_required_columns(
            file_path=file_path,
            required_column_candidates=[
                ["Client Name"],
                ["Contract Number"],
                ["Total Commission"],
                ["Planner"],
            ],
            header=0,
            exclude_sheet_names=[],
        )

        print(f"   source_sheet_name: {source_sheet_name}")
        print(f"   raw shape: {df.shape if isinstance(df, pd.DataFrame) else None}")

        if df is None or df.empty:
            print("🔴 No usable Sanlam sheet could be identified.")
            return None

        # -------------------------------------------------------------------
        # Step 2: Detect source columns
        # -------------------------------------------------------------------
        print("🟦 Step 2: Detect source columns")
        client_name_col = find_matching_column(df.columns, ["Client Name"])
        contract_col = find_matching_column(df.columns, ["Contract Number"])
        total_commission_col = find_matching_column(df.columns, ["Total Commission"])
        planner_col = find_matching_column(df.columns, ["Planner"])

        print(f"   client_name_col       : {client_name_col}")
        print(f"   contract_col          : {contract_col}")
        print(f"   total_commission_col  : {total_commission_col}")
        print(f"   planner_col           : {planner_col}")

        if contract_col is None or total_commission_col is None:
            print("🔴 Required Sanlam columns could not be identified.")
            return None

        # -------------------------------------------------------------------
        # Step 3: Build normalized output df
        # -------------------------------------------------------------------
        print("🟦 Step 3: Build normalized output df")
        normalized_df = pd.DataFrame()

        if client_name_col:
            normalized_df["client_name"] = df[client_name_col].apply(normalize_text_preserve_case)
        else:
            normalized_df["client_name"] = ""

        normalized_df["product_house"] = "sanlam"
        normalized_df["commission_month"] = comm_month
        normalized_df["contract_number"] = df[contract_col].apply(normalize_text_preserve_case)
        normalized_df["total_commission"] = coerce_series_to_numeric(df[total_commission_col])

        if planner_col:
            normalized_df["planner"] = (
                df[planner_col]
                .apply(normalize_text_preserve_case)
                .replace("", "UNIFY (PTY) LTD")
            )
        else:
            normalized_df["planner"] = "UNIFY (PTY) LTD"

        # -------------------------------------------------------------------
        # Step 4: Filter invalid rows
        # -------------------------------------------------------------------
        print("🟦 Step 4: Filter invalid rows")
        normalized_df = normalized_df[
            normalized_df["contract_number"].apply(lambda x: normalize_text(x) != "")
        ].copy()

        normalized_df = normalized_df[
            normalized_df["total_commission"].notna()
        ].copy()

        normalized_df = normalized_df.reset_index(drop=True)
        print(f"   normalized_df shape after filters: {normalized_df.shape}")

        if normalized_df.empty:
            print("🔴 No usable Sanlam rows remain after filtering.")
            return None

        # -------------------------------------------------------------------
        # Step 5: Reorder final columns
        # -------------------------------------------------------------------
        print("🟦 Step 5: Reorder final columns")
        normalized_df = normalized_df[
            [
                "client_name",
                "product_house",
                "commission_month",
                "contract_number",
                "total_commission",
                "planner",
            ]
        ].copy()

        print("✅ Sanlam normalized read complete")
        print(normalized_df.head(10))

        return normalized_df

    except Exception as e:
        print(f"🔴 Error in read_sanlam_df: {e}")
        print(traceback.format_exc())
        return None


# ---------------------------------------------------------------------------
# 5.4.8 Reader: Sirago
# ---------------------------------------------------------------------------
@register_reader("Sirago")
def read_sirago_df(file_path, **kwargs):
    """
    Fallback reader for Sirago files based on the uploaded example.

    Important:
    - the uploaded Sirago reference workbook only contained an already-
      normalized processed_data-style sheet
    - because of that, this reader currently re-imports a normalized sheet
      rather than parsing a raw Sirago statement layout
    """
    try:
        print("------------------------------------------------------------")
        print("🧾 5.4.8 Reader: read_sirago_df")
        print(f"   file_path: {file_path}")

        # -------------------------------------------------------------------
        # Step 0: Get passed variables
        # -------------------------------------------------------------------
        print("🟦 Step 0: Get passed variables")
        comm_month = kwargs.get("comm_month")
        comm_tables_main_df = kwargs.get("comm_tables_main_df")

        print(f"   comm_month: {comm_month}")
        print(
            f"   comm_tables_main_df shape: "
            f"{comm_tables_main_df.shape if isinstance(comm_tables_main_df, pd.DataFrame) else None}"
        )

        # -------------------------------------------------------------------
        # Step 1: Read best matching normalized sheet
        # -------------------------------------------------------------------
        print("🟦 Step 1: Read best matching normalized sheet")
        source_sheet_name, df = read_best_matching_sheet_by_required_columns(
            file_path=file_path,
            required_column_candidates=[
                ["Client Name"],
                ["Contract Number"],
                ["Total Commission"],
                ["Planner"],
            ],
            header=0,
            exclude_sheet_names=[],
        )

        print(f"   source_sheet_name: {source_sheet_name}")
        print(f"   raw shape: {df.shape if isinstance(df, pd.DataFrame) else None}")

        if df is None or df.empty:
            print("🔴 No usable Sirago sheet could be identified.")
            return None

        # -------------------------------------------------------------------
        # Step 2: Detect source columns
        # -------------------------------------------------------------------
        print("🟦 Step 2: Detect source columns")
        client_name_col = find_matching_column(df.columns, ["Client Name"])
        contract_col = find_matching_column(df.columns, ["Contract Number"])
        total_commission_col = find_matching_column(df.columns, ["Total Commission"])
        planner_col = find_matching_column(df.columns, ["Planner"])

        print(f"   client_name_col       : {client_name_col}")
        print(f"   contract_col          : {contract_col}")
        print(f"   total_commission_col  : {total_commission_col}")
        print(f"   planner_col           : {planner_col}")

        if contract_col is None or total_commission_col is None:
            print("🔴 Required Sirago columns could not be identified.")
            return None

        # -------------------------------------------------------------------
        # Step 3: Build normalized output df
        # -------------------------------------------------------------------
        print("🟦 Step 3: Build normalized output df")
        normalized_df = pd.DataFrame()

        if client_name_col:
            normalized_df["client_name"] = df[client_name_col].apply(normalize_text_preserve_case)
        else:
            normalized_df["client_name"] = ""

        normalized_df["product_house"] = "sirago"
        normalized_df["commission_month"] = comm_month
        normalized_df["contract_number"] = df[contract_col].apply(normalize_text_preserve_case)
        normalized_df["total_commission"] = coerce_series_to_numeric(df[total_commission_col])

        if planner_col:
            normalized_df["planner"] = (
                df[planner_col]
                .apply(normalize_text_preserve_case)
                .replace("", "UNIFY (PTY) LTD")
            )
        else:
            normalized_df["planner"] = "UNIFY (PTY) LTD"

        # -------------------------------------------------------------------
        # Step 4: Filter invalid rows
        # -------------------------------------------------------------------
        print("🟦 Step 4: Filter invalid rows")
        normalized_df = normalized_df[
            normalized_df["contract_number"].apply(lambda x: normalize_text(x) != "")
        ].copy()

        normalized_df = normalized_df[
            normalized_df["total_commission"].notna()
        ].copy()

        normalized_df = normalized_df.reset_index(drop=True)
        print(f"   normalized_df shape after filters: {normalized_df.shape}")

        if normalized_df.empty:
            print("🔴 No usable Sirago rows remain after filtering.")
            return None

        # -------------------------------------------------------------------
        # Step 5: Reorder final columns
        # -------------------------------------------------------------------
        print("🟦 Step 5: Reorder final columns")
        normalized_df = normalized_df[
            [
                "client_name",
                "product_house",
                "commission_month",
                "contract_number",
                "total_commission",
                "planner",
            ]
        ].copy()

        print("✅ Sirago normalized read complete")
        print(normalized_df.head(10))

        return normalized_df

    except Exception as e:
        print(f"🔴 Error in read_sirago_df: {e}")
        print(traceback.format_exc())
        return None






























































































































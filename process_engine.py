import os
import re
import traceback
import pandas as pd
import xlwings as xw

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
    Return a copy of df with cleaned column headers.

    Purpose:
    - makes header matching more reliable
    - removes line breaks
    - trims spaces
    - collapses repeated spaces
    """
    df = df.copy()

    cleaned_columns = []
    for col in df.columns:
        cleaned_col = normalize_text_preserve_case(col)
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
# 1.3 File Reading Utilities
# ---------------------------------------------------------------------------
# ---------------------------------------------------------------------------
# 1.3.1 Utility: Read Raw First Sheet
# ---------------------------------------------------------------------------
def read_raw_first_sheet(file_path, header=None):
    """
    Read the first sheet of an Excel file using pandas.
    """
    print(f"   [UTILITY] Reading raw first sheet: {file_path}")
    return pd.read_excel(file_path, sheet_name=0, header=header)

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
                "Excel Files (*.xlsx *.xls *.xlsm)"
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
            "Momentum_MFP", "MUA", "Nedgroup", "Old_Mutual_Short_Term",
            "Old_Mutual_Life_Invest", "Sanlam", "Santam", "SAU", "Sirago",
            "Sirago_Silk", "Stanlib", "Stratum", "Stratum_Silk", "Turnberry",
            "Zestlife"
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
    """
    Official version write logic.

    Purpose:
    - Open the selected commission workbook file
    - Create or re-use a sheet called 'processed_data'
    - Clear the old contents
    - Write the processed / normalized DF to that sheet
    - Save and close the selected workbook

    Notes:
    - Does NOT change DF logic
    - Does NOT write to Comm_Tables
    - Writes headers
    - Does NOT write index column
    """
    try:
        print("------------------------------------------------------------")
        print("🟦 5.1.2 Write Official DF To Selected File")
        print(f"   file_path: {file_path}")

        if df is None or not isinstance(df, pd.DataFrame) or df.empty:
            print("🟠 DF is empty or invalid. Nothing to write.")
            return "no_data"

        print(f"   df shape: {df.shape}")
        print(f"   df columns: {list(df.columns)}")

        # -------------------------------------------------------------------
        # Step 1: Open selected workbook
        # -------------------------------------------------------------------
        print("🟦 Step 1: Open selected workbook")
        target_wb = xw.Book(file_path)
        print(f"   opened workbook: {target_wb.name}")

        # -------------------------------------------------------------------
        # Step 2: Get or create processed_data sheet
        # -------------------------------------------------------------------
        print("🟦 Step 2: Get or create 'processed_data' sheet")
        processed_sheet_name = "processed_data"

        existing_sheet_names = [s.name.lower() for s in target_wb.sheets]

        if processed_sheet_name.lower() in existing_sheet_names:
            sh_processed_data = None
            for s in target_wb.sheets:
                if s.name.strip().lower() == processed_sheet_name.lower():
                    sh_processed_data = s
                    break
            print(f"   existing sheet found: {sh_processed_data.name}")
        else:
            sh_processed_data = target_wb.sheets.add(processed_sheet_name)
            print(f"   created new sheet: {sh_processed_data.name}")

        # -------------------------------------------------------------------
        # Step 3: Clear old contents
        # -------------------------------------------------------------------
        print("🟦 Step 3: Clear old contents on processed_data")
        sh_processed_data.range("A1:ZZ4000").clear_contents()

        # -------------------------------------------------------------------
        # Step 4: Write DF
        # -------------------------------------------------------------------
        print("🟦 Step 4: Write DF to processed_data!A1")
        sh_processed_data.range("A1").options(index=False, header=True).value = df

        try:
            sh_processed_data.autofit()
            print("   autofit applied")
        except Exception as autofit_error:
            print(f"   ⚠️ Autofit skipped: {autofit_error}")

        # -------------------------------------------------------------------
        # Step 5: Save and close workbook
        # -------------------------------------------------------------------
        print("🟦 Step 5: Save and close selected workbook")
        target_wb.save()
        target_wb.close()

        print("✅ Official DF write complete")
        return "success"

    except Exception as e:
        print(f"🔴 Error in f5_1_2_write_official_df_to_selected_file: {e}")
        print(traceback.format_exc())
        return "failed"


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
# 5.3 Range-Based / Openpyxl Reader Functions
# ---------------------------------------------------------------------------

# ---------------------------------------------------------------------------
# 5.4 Special-Layout Reader Functions
# ---------------------------------------------------------------------------































































































































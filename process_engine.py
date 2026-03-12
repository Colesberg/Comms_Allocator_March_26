import os
import re
import traceback
import pandas as pd

from PyQt5.QtWidgets import (
    QApplication, QWidget, QDialog, QLabel, QPushButton, QVBoxLayout,
    QHBoxLayout, QFileDialog, QMessageBox, QListWidget, QComboBox,
    QTextEdit, QRadioButton, QButtonGroup
)
from PyQt5.QtCore import Qt
import sys

# ---------------------------------------------------------------------------
# F1: Confirm Test / Official / File / Folder
# ---------------------------------------------------------------------------
def f1_confirm_test_file_folder():
    """
    Launches a PyQt5 widget to confirm:
      - Test or Official
      - If Official: File or Folder
      - If Official + File: Clear table?
      - If Official + Folder: clear_table defaults to 'yes'
      - If Test: file_folder defaults to 'file', clear_table = None

    Returns:
      test_official, file_folder, clear_table
    """
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

    if owns_app:
        app.quit()

    return test_official, file_folder, clear_table


# ---------------------------------------------------------------------------
# Widget: Confirm Run Type / File / Folder / Clear Table
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

        # Test / Official
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
            # test => defaults to file, clear_table irrelevant
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
                # official + folder => clear_table defaults to yes
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
# F3A Widget: Prompt User To Pick File Or Folder
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

        # -------------------------------------------------------------------
        # Title
        # -------------------------------------------------------------------
        title_label = QLabel("Select the required input")
        title_label.setAlignment(Qt.AlignCenter)
        title_label.setStyleSheet("font-size: 16px; font-weight: bold;")
        main_layout.addWidget(title_label)

        # -------------------------------------------------------------------
        # Mode label
        # -------------------------------------------------------------------
        if self.file_folder == "folder":
            mode_text = "Mode selected: Folder import"
        else:
            mode_text = "Mode selected: File import"

        self.mode_label = QLabel(mode_text)
        self.mode_label.setAlignment(Qt.AlignCenter)
        self.mode_label.setStyleSheet("font-size: 13px;")
        main_layout.addWidget(self.mode_label)

        # -------------------------------------------------------------------
        # Selected path display
        # -------------------------------------------------------------------
        self.path_label = QLabel("No path selected yet.")
        self.path_label.setWordWrap(True)
        self.path_label.setStyleSheet(
            "border: 1px solid #999; padding: 8px; background-color: #f7f7f7;"
        )
        main_layout.addWidget(self.path_label)

        # -------------------------------------------------------------------
        # Browse button row
        # -------------------------------------------------------------------
        browse_layout = QHBoxLayout()

        self.browse_button = QPushButton("Browse")
        self.browse_button.clicked.connect(self.browse_path)
        browse_layout.addWidget(self.browse_button)

        main_layout.addLayout(browse_layout)

        # -------------------------------------------------------------------
        # Continue / Exit buttons
        # -------------------------------------------------------------------
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
                "Select Excel File",
                "",
                "Excel Files (*.xlsx *.xls *.xlsm *.xlsb)"
            )
            if file_path:
                self.selected_path = file_path
                self.path_label.setText(file_path)

    def confirm_selection(self):
        if not self.selected_path:
            QMessageBox.warning(
                self,
                "No Selection",
                "Please select a file or folder before continuing."
            )
            return

        self.accept()
        
# ---------------------------------------------------------------------------
# F3A: Prompt User To Pick File Or Folder
# ---------------------------------------------------------------------------
def f3a_prompt_user_to_pick_file_or_folder(file_folder):
    """
    Launch a PyQt5 widget to let the user select either:
    - a single Excel file
    - a folder containing Excel files

    Returns:
        selected_path (str) if confirmed
        None if cancelled
    """
    print("   [F3A] Launching file/folder picker widget...")
    print(f"   [F3A] file_folder mode received: {file_folder}")

    app = QApplication.instance()
    app_created_here = False

    if app is None:
        app = QApplication([])
        app_created_here = True

    dialog = FileFolderPickerWidget(file_folder=file_folder)
    result = dialog.exec_()

    selected_path = dialog.selected_path if result == QDialog.Accepted else None

    print(f"   [F3A] selected_path: {selected_path}")

    if app_created_here:
        app.quit()

    return selected_path

      # ---------------------------------------------------------------------------

# Reader Registry
# ---------------------------------------------------------------------------
READER_REGISTRY = {} 

# ---------------------------------------------------------------------------
# Utility: Register Reader
# ---------------------------------------------------------------------------
def register_reader(comm_type):
    """
    Decorator to register a reader function against a comm_type.
    """
    def decorator(func):
        READER_REGISTRY[comm_type] = func
        return func
    return decorator     

# ---------------------------------------------------------------------------
# Utility: Get Reader For Comm Type
# ---------------------------------------------------------------------------
def get_reader_for_comm_type(comm_type):
    reader_func = READER_REGISTRY.get(comm_type)
    print(f"   [ENGINE] get_reader_for_comm_type('{comm_type}') -> {reader_func}")
    return reader_func

# ---------------------------------------------------------------------------
# Utility: Normalize Text
# ---------------------------------------------------------------------------
def normalize_text(value):
    if pd.isna(value):
        return ""
    return str(value).strip().lower()

# ---------------------------------------------------------------------------
# Utility: Find Row Index By First Column Value
# ---------------------------------------------------------------------------
def find_row_index_by_first_column_value(df, target_text):
    """
    Finds the first row index where column 0 equals target_text after normalization.
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
# Utility: Set Header From Row
# ---------------------------------------------------------------------------
def set_header_from_row(df, header_row_idx):
    """
    Use a row as the dataframe header, then remove rows above it.
    """
    df = df.copy()
    df.columns = df.iloc[header_row_idx]
    df = df.drop(df.index[:header_row_idx + 1])
    df = df.reset_index(drop=True)
    return df

# ---------------------------------------------------------------------------
# Utility: Drop Blank Rows By Column Position
# ---------------------------------------------------------------------------
def drop_blank_rows_by_column_position(df, column_position):
    """
    Drops rows where the specified column position is blank / NaN.
    """
    df = df.copy()

    if column_position >= len(df.columns):
        print(f"🟠 column_position {column_position} is outside df columns.")
        return df

    target_col = df.columns[column_position]
    df = df.dropna(subset=[target_col])

    return df

# ---------------------------------------------------------------------------
# Utility: Read Raw First Sheet
# ---------------------------------------------------------------------------
def read_raw_first_sheet(file_path, header=None):
    print(f"   [UTILITY] Reading raw first sheet: {file_path}")
    return pd.read_excel(file_path, sheet_name=0, header=header)


            
# ---------------------------------------------------------------------------
# Widget: Prompt Comm Type Selection
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
# Widget: Processing Error
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
        

# ---------------------------------------------------------------------------
# Reader: AGBus
# ---------------------------------------------------------------------------
@register_reader("AGBus")
def read_agbus_df(file_path, **kwargs):
    try:
        print("------------------------------------------------------------")
        print("🧾 Reader: read_agbus_df")
        print(f"   file_path: {file_path}")

        # Step 1: Read raw file
        print("🟦 Step 1: Read raw first sheet")
        df = read_raw_first_sheet(file_path, header=None)
        print(f"   raw shape: {df.shape}")

        # Step 2: Find 'Earning Year' row
        print("🟦 Step 2: Find 'Earning Year' row")
        header_row_idx = find_row_index_by_first_column_value(df, "Earning Year")
        print(f"   header_row_idx: {header_row_idx}")

        if header_row_idx is None:
            print("🔴 'Earning Year' row not found.")
            return None

        # Step 3: Set header from row
        print("🟦 Step 3: Set header from target row")
        df = set_header_from_row(df, header_row_idx)
        print(f"   shape after header set: {df.shape}")

        # Step 4: Drop blank rows by second column
        print("🟦 Step 4: Drop blank rows from second column")
        df = drop_blank_rows_by_column_position(df, 1)
        print(f"   final shape: {df.shape}")

        return df

    except Exception as e:
        print(f"🔴 Error in read_agbus_df: {e}")
        print(traceback.format_exc())
        return None

# ---------------------------------------------------------------------------
# F4.2.1: Detect Comm Type From Filename
# ---------------------------------------------------------------------------
def f4_2_1_detect_comm_type_label(file_path):
    file_name = os.path.basename(file_path).lower().strip()

    print("------------------------------------------------------------")
    print("🟦 Detect comm type from filename")
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
        "mfp": "Momentum_MFP",
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

    for key_text, comm_type in detection_map.items():
        if key_text in file_name:
            print(f"✅ Matched '{key_text}' -> '{comm_type}'")
            return comm_type

    print("🟠 No comm type detected from filename.")
    return None


# ---------------------------------------------------------------------------
# F4.2.2: Prompt Comm Type Only If Needed
# ---------------------------------------------------------------------------
def f4_2_2_prompt_comm_type_only_if_needed(detected_comm_type=None):
    print("   [ENGINE] Launching comm type selection widget...")
    print(f"   [ENGINE] detected_comm_type: {detected_comm_type}")

    app = QApplication.instance()
    app_created_here = False

    if app is None:
        app = QApplication([])
        app_created_here = True

    dialog = CommTypeSelectionWidget(detected_comm_type=detected_comm_type)
    result = dialog.exec_()

    selected_comm_type = dialog.selected_comm_type if result == QDialog.Accepted else None

    print(f"   [ENGINE] selected_comm_type: {selected_comm_type}")

    if app_created_here:
        app.quit()

    return selected_comm_type


# ---------------------------------------------------------------------------
# F4.4: Run Engineered Reader
# ---------------------------------------------------------------------------
def f4_4_run_engineered_reader(comm_type, file_path, comm_month, comm_tables_main_df, reader_func=None):
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


# ---------------------------------------------------------------------------
# F4.5: Show Processing Error Widget
# ---------------------------------------------------------------------------
def f4_5_show_processing_error_widget(file_path, error_message):
    print("   [ENGINE] Showing processing error widget...")
    print(f"   [ENGINE] file_path    : {file_path}")
    print(f"   [ENGINE] error_message: {error_message}")

    app = QApplication.instance()
    app_created_here = False

    if app is None:
        app = QApplication([])
        app_created_here = True

    dialog = ProcessingErrorWidget(file_path=file_path, error_message=error_message)
    dialog.exec_()

    if app_created_here:
        app.quit()
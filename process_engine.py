import sys

from PyQt5.QtWidgets import (
    QApplication,
    QWidget,
    QLabel,
    QPushButton,
    QVBoxLayout,
    QHBoxLayout,
    QComboBox,
    QMessageBox
)
from PyQt5.QtCore import Qt


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
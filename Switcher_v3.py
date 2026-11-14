import sys
import json
import subprocess
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook, Workbook

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QApplication,
    QComboBox,
    QDialog,
    QFormLayout,
    QFrame,
    QHBoxLayout,
    QInputDialog,
    QLabel,
    QListView,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QVBoxLayout,
    QWidget,
    QListWidget,
    QLineEdit,
)


# ============================================================
# MODERN DARK THEME STYLESHEET
# ============================================================

APP_STYLESHEET = """
QWidget {
    background-color: #12131c;
    color: #e2e8f0;
    font-family: 'Segoe UI', system-ui, -apple-system, sans-serif;
    font-size: 13px;
}

QMainWindow, QDialog {
    background-color: #12131c;
}

/* Card Container */
QFrame#cardFrame {
    background-color: #1a1c2b;
    border: 1px solid #2d3148;
    border-radius: 10px;
    padding: 14px;
}

/* Labels */
QLabel {
    color: #cbd5e1;
    font-size: 13px;
}

/* Inputs & Comboboxes */
QLineEdit, QComboBox {
    background-color: #222638;
    border: 1px solid #363b57;
    border-radius: 7px;
    padding: 7px 12px;
    color: #f8fafc;
    font-size: 13px;
    min-height: 22px;
}

QLineEdit:focus, QComboBox:focus {
    border: 1.5px solid #6366f1;
    background-color: #272c40;
}

/* Combobox Arrow */
QComboBox::drop-down {
    subcontrol-origin: padding;
    subcontrol-position: top right;
    width: 28px;
    border-left: 1px solid #363b57;
    border-top-right-radius: 7px;
    border-bottom-right-radius: 7px;
}

QComboBox::down-arrow {
    width: 0;
    height: 0;
    border-left: 5px solid transparent;
    border-right: 5px solid transparent;
    border-top: 6px solid #94a3b8;
}

QComboBox::down-arrow:hover {
    border-top-color: #6366f1;
}

/* Combobox Popup View (Ensures all items are crisp & visible) */
QComboBox QAbstractItemView {
    background-color: #1a1c2b;
    color: #f8fafc;
    border: 1px solid #3b4263;
    border-radius: 8px;
    padding: 4px;
    selection-background-color: #6366f1;
    selection-color: #ffffff;
    outline: none;
}

QComboBox QAbstractItemView::item {
    min-height: 28px;
    padding: 6px 10px;
    border-radius: 4px;
    color: #f8fafc;
}

QComboBox QAbstractItemView::item:hover {
    background-color: #2b3047;
    color: #ffffff;
}

QComboBox QAbstractItemView::item:selected {
    background-color: #6366f1;
    color: #ffffff;
}

/* List Widgets */
QListWidget {
    background-color: #1a1c2b;
    border: 1px solid #2d3148;
    border-radius: 8px;
    padding: 6px;
    color: #f8fafc;
}

QListWidget::item {
    padding: 7px 10px;
    border-radius: 5px;
    margin-bottom: 2px;
}

QListWidget::item:hover {
    background-color: #272c40;
}

QListWidget::item:selected {
    background-color: #6366f1;
    color: #ffffff;
}

/* Default Buttons */
QPushButton {
    background-color: #25293d;
    color: #f1f5f9;
    border: 1px solid #383e5c;
    border-radius: 7px;
    padding: 8px 14px;
    font-weight: 600;
    font-size: 13px;
}

QPushButton:hover {
    background-color: #313652;
    border-color: #4f5780;
}

QPushButton:pressed {
    background-color: #1c1f2e;
}

/* Apply / Primary Accent Button */
QPushButton#applyBtn {
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #6366f1, stop:1 #4f46e5);
    color: #ffffff;
    border: none;
    border-radius: 8px;
    padding: 10px 20px;
    font-size: 13px;
    font-weight: bold;
    letter-spacing: 0.5px;
}

QPushButton#applyBtn:hover {
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #7579f6, stop:1 #5b53ed);
}

QPushButton#applyBtn:pressed {
    background: #4338ca;
}

/* Secondary Button */
QPushButton#secondaryBtn {
    background-color: #1e2233;
    border: 1px solid #353a54;
    color: #cbd5e1;
}

QPushButton#secondaryBtn:hover {
    background-color: #292e45;
    border-color: #4b5275;
    color: #ffffff;
}

/* Danger Button (Delete) */
QPushButton#dangerBtn {
    background-color: #2d1b22;
    border: 1px solid #572530;
    color: #f87171;
}

QPushButton#dangerBtn:hover {
    background-color: #ef4444;
    border-color: #dc2626;
    color: #ffffff;
}

/* Success Button (Add) */
QPushButton#successBtn {
    background-color: #182823;
    border: 1px solid #234d3f;
    color: #34d399;
}

QPushButton#successBtn:hover {
    background-color: #10b981;
    border-color: #059669;
    color: #ffffff;
}

/* Scrollbars */
QScrollBar:vertical {
    background: #12131c;
    width: 10px;
    border-radius: 5px;
}

QScrollBar::handle:vertical {
    background: #2d3148;
    border-radius: 5px;
    min-height: 25px;
}

QScrollBar::handle:vertical:hover {
    background: #454b6e;
}

QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
    height: 0px;
}

/* Tooltips */
QToolTip {
    background-color: #1a1c2b;
    color: #f8fafc;
    border: 1px solid #454b6e;
    border-radius: 5px;
    padding: 5px 8px;
    font-size: 12px;
}
"""


# ============================================================
# FILE PATHS
# ============================================================

SETTINGS_FILE = Path(
    r"C:\Users\Nishant\.claude\settings.json"
)

KEY_FILE = Path(
    r"C:\My Space\Data\keys.xlsx"
)

PROVIDERS_FILE = Path(__file__).with_name(
    "providers.json"
)

SWITCH_HISTORY_FILE = Path(
    r"C:\My Space\Data\model_switch_history.xlsx"
)


# ============================================================
# CLAUDE CODE SETTINGS
# ============================================================

MODEL_SETTINGS = [
    "ANTHROPIC_MODEL",
    "ANTHROPIC_SMALL_FAST_MODEL",
    "ANTHROPIC_DEFAULT_SONNET_MODEL",
    "ANTHROPIC_DEFAULT_OPUS_MODEL",
    "ANTHROPIC_DEFAULT_HAIKU_MODEL",
]




# ============================================================
# PROVIDER JSON FUNCTIONS
# ============================================================

DEFAULT_PROVIDERS = {
    "providers": {
        "Ollama": {
            "base_url": "https://ollama.com",
            "models": ["gemma4:31b", "gpt-oss:120b", "gpt-oss:20b"]
        },
        "Openrouter": {
            "base_url": "https://openrouter.ai/api",
            "models": ["nvidia/nemotron-3-ultra-550b-a55b:free", "qwen/qwen3.8-27b:free"]
        },
        "groq": {
            "base_url": "https://api.groq.com/openai/v1",
            "models": ["openai/gpt-oss-20b", "llama-3.3-70b-versatile", "llama-3.1-8b-instant"]
        }
    }
}

def ensure_provider_file():

    if not PROVIDERS_FILE.exists():

        PROVIDERS_FILE.write_text(
            json.dumps(
                DEFAULT_PROVIDERS,
                indent=2,
                ensure_ascii=False
            ),
            encoding="utf-8"
        )


def load_providers():

    ensure_provider_file()

    try:

        data = json.loads(
            PROVIDERS_FILE.read_text(
                encoding="utf-8"
            )
        )

        if not isinstance(
            data.get("providers"),
            dict
        ):

            raise ValueError(
                "Invalid providers.json format."
            )

        return data

    except Exception as exc:

        raise RuntimeError(
            f"Could not read providers.json:\n{exc}"
        )


def save_providers(data):

    PROVIDERS_FILE.write_text(
        json.dumps(
            data,
            indent=2,
            ensure_ascii=False
        ),
        encoding="utf-8"
    )


# ============================================================
# EXCEL FUNCTIONS
# ============================================================

def ensure_key_file():

    if not KEY_FILE.exists():

        KEY_FILE.parent.mkdir(
            parents=True,
            exist_ok=True
        )

        wb = Workbook()

        ws = wb.active

        ws.title = "Keys"

        ws.append([
            "Provider",
            "Nickname",
            "API Key"
        ])

        wb.save(
            KEY_FILE
        )

        wb.close()


def save_switch_history(
    nickname,
    model
):

    now = datetime.now()

    if SWITCH_HISTORY_FILE.exists():

        wb = load_workbook(
            SWITCH_HISTORY_FILE
        )

        ws = wb.active

    else:

        SWITCH_HISTORY_FILE.parent.mkdir(
            parents=True,
            exist_ok=True
        )

        wb = Workbook()

        ws = wb.active

        ws.title = "Switch History"

        ws.append([
            "Date",
            "Time",
            "Nickname",
            "Model"
        ])

    ws.append([
        now.strftime("%Y-%m-%d"),
        now.strftime("%H:%M:%S"),
        nickname,
        model
    ])

    wb.save(
        SWITCH_HISTORY_FILE
    )

    wb.close()


# ============================================================
# PROVIDER SETTINGS DIALOG
# ============================================================

class ProviderSettingsDialog(QDialog):

    def __init__(self, parent, providers):
        super().__init__(parent)

        self.parent_window = parent
        self.providers = providers

        self.setWindowTitle("⚙ Provider & Model Settings")
        self.resize(700, 520)

        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(18, 18, 18, 18)
        main_layout.setSpacing(14)

        # ----------------------------------------------------
        # CARD CONTAINER
        # ----------------------------------------------------
        card = QFrame()
        card.setObjectName("cardFrame")
        card_layout = QVBoxLayout(card)
        card_layout.setSpacing(12)

        form = QFormLayout()
        form.setSpacing(10)

        # Provider selection row with Add/Delete buttons
        provider_row = QHBoxLayout()
        self.provider_dropdown = QComboBox()
        self.provider_dropdown.setView(QListView())
        self.provider_dropdown.addItems(self.providers.get("providers", {}).keys())
        self.provider_dropdown.currentTextChanged.connect(self.load_provider)

        add_prov_btn = QPushButton("➕ Add Provider")
        add_prov_btn.setObjectName("successBtn")
        add_prov_btn.setToolTip("Add a new provider to providers.json")
        add_prov_btn.clicked.connect(self.add_provider)

        del_prov_btn = QPushButton("🗑 Delete")
        del_prov_btn.setObjectName("dangerBtn")
        del_prov_btn.setToolTip("Delete selected provider from providers.json")
        del_prov_btn.clicked.connect(self.delete_provider)

        provider_row.addWidget(self.provider_dropdown, stretch=2)
        provider_row.addWidget(add_prov_btn)
        provider_row.addWidget(del_prov_btn)

        form.addRow("Provider:", provider_row)

        # Base URL
        self.url_edit = QLineEdit()
        self.url_edit.setPlaceholderText("https://api.example.com/v1")
        form.addRow("Base URL:", self.url_edit)

        card_layout.addLayout(form)
        main_layout.addWidget(card)

        # ----------------------------------------------------
        # MODELS SECTION
        # ----------------------------------------------------
        models_header = QHBoxLayout()
        self.models_count_lbl = QLabel("Models (0):")
        self.models_count_lbl.setStyleSheet("font-weight: bold; color: #cbd5e1; font-size: 13px;")

        add_model_btn = QPushButton("➕ Add Model")
        add_model_btn.setObjectName("secondaryBtn")
        add_model_btn.clicked.connect(self.add_model)

        delete_model_btn = QPushButton("🗑 Delete Model")
        delete_model_btn.setObjectName("dangerBtn")
        delete_model_btn.clicked.connect(self.delete_model)

        models_header.addWidget(self.models_count_lbl)
        models_header.addStretch()
        models_header.addWidget(add_model_btn)
        models_header.addWidget(delete_model_btn)

        main_layout.addLayout(models_header)

        self.model_list = QListWidget()
        main_layout.addWidget(self.model_list, stretch=1)

        # ----------------------------------------------------
        # ACTION BUTTONS
        # ----------------------------------------------------
        bottom_buttons = QHBoxLayout()
        save_btn = QPushButton("💾 SAVE CHANGES")
        save_btn.setObjectName("applyBtn")
        save_btn.setMinimumHeight(40)
        save_btn.clicked.connect(self.save)

        close_btn = QPushButton("Close")
        close_btn.setObjectName("secondaryBtn")
        close_btn.setMinimumHeight(40)
        close_btn.clicked.connect(self.reject)

        bottom_buttons.addStretch()
        bottom_buttons.addWidget(save_btn)
        bottom_buttons.addWidget(close_btn)

        main_layout.addLayout(bottom_buttons)

        # Load initial provider
        if self.provider_dropdown.count():
            self.load_provider(self.provider_dropdown.currentText())

    def load_provider(self, provider):
        data = self.providers.get("providers", {}).get(provider, {})
        self.url_edit.setText(data.get("base_url", ""))
        self.model_list.clear()
        models = data.get("models", [])
        self.model_list.addItems(models)
        self.models_count_lbl.setText(f"Models ({len(models)}):")

    def add_provider(self):
        name, ok = QInputDialog.getText(
            self,
            "Add Provider",
            "Enter new provider name (e.g. DeepSeek, Together, Anthropic):"
        )
        if not ok:
            return
        name = name.strip()
        if not name:
            QMessageBox.warning(self, "Invalid Name", "Provider name cannot be empty.")
            return

        providers_dict = self.providers.setdefault("providers", {})
        for existing in providers_dict.keys():
            if existing.lower() == name.lower():
                QMessageBox.warning(
                    self,
                    "Provider Exists",
                    f"A provider named '{existing}' already exists."
                )
                self.provider_dropdown.setCurrentText(existing)
                return

        # Initialize provider entry
        providers_dict[name] = {
            "base_url": "",
            "models": []
        }

        self.provider_dropdown.blockSignals(True)
        self.provider_dropdown.addItem(name)
        self.provider_dropdown.setCurrentText(name)
        self.provider_dropdown.blockSignals(False)

        self.load_provider(name)
        self.url_edit.setFocus()

        QMessageBox.information(
            self,
            "Provider Added",
            f"Provider '{name}' added!\n\nPlease enter its Base URL, add model(s), and click 'SAVE CHANGES'."
        )

    def delete_provider(self):
        provider = self.provider_dropdown.currentText().strip()
        if not provider:
            return

        reply = QMessageBox.question(
            self,
            "Delete Provider",
            f"Are you sure you want to delete '{provider}' and all its models from providers.json?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No
        )
        if reply != QMessageBox.StandardButton.Yes:
            return

        providers_dict = self.providers.setdefault("providers", {})
        if provider in providers_dict:
            del providers_dict[provider]

        try:
            save_providers(self.providers)
            self.parent_window.load_data()
        except Exception as exc:
            QMessageBox.critical(self, "Error", f"Failed to save changes:\n{exc}")
            return

        idx = self.provider_dropdown.findText(provider)
        if idx >= 0:
            self.provider_dropdown.removeItem(idx)

        if self.provider_dropdown.count() > 0:
            self.load_provider(self.provider_dropdown.currentText())
        else:
            self.url_edit.clear()
            self.model_list.clear()
            self.models_count_lbl.setText("Models (0):")

        QMessageBox.information(self, "Deleted", f"Provider '{provider}' was removed.")

    def add_model(self):
        model, ok = QInputDialog.getText(
            self,
            "Add Model",
            "Enter Model ID (e.g. llama-3.3-70b-versatile, gpt-4o):"
        )
        if not ok:
            return
        model = model.strip()
        if not model:
            return

        existing = [self.model_list.item(i).text() for i in range(self.model_list.count())]
        if model in existing:
            QMessageBox.warning(self, "Duplicate", "This model already exists in the list.")
            return

        self.model_list.addItem(model)
        self.models_count_lbl.setText(f"Models ({self.model_list.count()}):")

    def delete_model(self):
        item = self.model_list.currentItem()
        if not item:
            QMessageBox.warning(self, "Delete Model", "Please select a model from the list first.")
            return

        self.model_list.takeItem(self.model_list.row(item))
        self.models_count_lbl.setText(f"Models ({self.model_list.count()}):")

    def save(self):
        provider = self.provider_dropdown.currentText().strip()
        if not provider:
            QMessageBox.warning(self, "No Provider", "Please select or add a provider first.")
            return

        base_url = self.url_edit.text().strip()
        models = [
            self.model_list.item(i).text().strip()
            for i in range(self.model_list.count())
            if self.model_list.item(i).text().strip()
        ]

        self.providers.setdefault("providers", {})[provider] = {
            "base_url": base_url,
            "models": models
        }

        try:
            save_providers(self.providers)
            self.parent_window.load_data()
            QMessageBox.information(
                self,
                "Saved",
                f"Provider '{provider}' configuration saved successfully."
            )
            self.accept()
        except Exception as exc:
            QMessageBox.critical(
                self,
                "Error",
                f"Could not save providers.json:\n{exc}"
            )


# ============================================================
# KEY MANAGER
# ============================================================

class KeyManagerDialog(QDialog):

    def __init__(self, parent):
        super().__init__(parent)

        self.parent_window = parent

        self.setWindowTitle("⚙ Key & Provider Management")
        self.resize(380, 290)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(18, 18, 18, 18)
        layout.setSpacing(10)

        header = QLabel("Manage Keys & Providers")
        header.setAlignment(Qt.AlignmentFlag.AlignCenter)
        header.setStyleSheet("font-size: 15px; font-weight: bold; color: #f8fafc; margin-bottom: 4px;")
        layout.addWidget(header)

        card = QFrame()
        card.setObjectName("cardFrame")
        card_layout = QVBoxLayout(card)
        card_layout.setSpacing(8)

        add_btn = QPushButton("➕  Add API Key")
        add_btn.setObjectName("successBtn")
        add_btn.setMinimumHeight(34)
        add_btn.clicked.connect(self.add_key)

        delete_btn = QPushButton("🗑  Delete Selected Key")
        delete_btn.setObjectName("dangerBtn")
        delete_btn.setMinimumHeight(34)
        delete_btn.clicked.connect(self.delete_key)

        provider_btn = QPushButton("⚙  Provider & Model Settings")
        provider_btn.setObjectName("secondaryBtn")
        provider_btn.setMinimumHeight(34)
        provider_btn.clicked.connect(self.provider_settings)

        tester_btn = QPushButton("🧪  Launch API Tester")
        tester_btn.setObjectName("secondaryBtn")
        tester_btn.setMinimumHeight(34)
        tester_btn.clicked.connect(self.open_api_tester)

        card_layout.addWidget(add_btn)
        card_layout.addWidget(delete_btn)
        card_layout.addWidget(provider_btn)
        card_layout.addWidget(tester_btn)

        layout.addWidget(card)

        close_btn = QPushButton("Close")
        close_btn.setMinimumHeight(32)
        close_btn.clicked.connect(self.accept)
        layout.addWidget(close_btn)


    # ========================================================
    # ADD KEY
    # ========================================================

    def add_key(self):

        providers = list(
            self.parent_window
            .providers[
                "providers"
            ]
            .keys()
        )

        if not providers:

            QMessageBox.warning(
                self,
                "No Providers",
                "No providers are configured."
            )

            return

        # ----------------------------------------------------
        # SELECT PROVIDER
        # ----------------------------------------------------

        provider, ok = QInputDialog.getItem(
            self,
            "Provider",
            "Select provider:",
            providers,
            0,
            False
        )

        if not ok:
            return

        # ----------------------------------------------------
        # NICKNAME
        # ----------------------------------------------------

        nickname, ok = QInputDialog.getText(
            self,
            "Add Key",
            "Nickname:"
        )

        if not ok:
            return

        nickname = nickname.strip()

        if not nickname:

            QMessageBox.warning(
                self,
                "Invalid",
                "Nickname cannot be empty."
            )

            return

        if nickname in self.parent_window.keys:

            QMessageBox.warning(
                self,
                "Duplicate",
                "This nickname already exists."
            )

            return

        # ----------------------------------------------------
        # API KEY
        # ----------------------------------------------------

        api_key, ok = QInputDialog.getText(
            self,
            "Add Key",
            "API Key:",
            QLineEdit.EchoMode.Password
        )

        if not ok:
            return

        api_key = api_key.strip()

        if not api_key:

            QMessageBox.warning(
                self,
                "Invalid",
                "API key cannot be empty."
            )

            return

        # ----------------------------------------------------
        # SAVE TO EXCEL
        # ----------------------------------------------------

        try:

            ensure_key_file()

            wb = load_workbook(
                KEY_FILE
            )

            ws = wb.active

            # Expected:
            #
            # Provider | Nickname | API Key

            ws.append([
                provider,
                nickname,
                api_key
            ])

            wb.save(
                KEY_FILE
            )

            wb.close()

            self.parent_window.load_data()

            QMessageBox.information(
                self,
                "Added",
                "API key added successfully."
            )

        except PermissionError:

            QMessageBox.critical(
                self,
                "Excel File Locked",
                "Close keys.xlsx in Excel and try again."
            )

        except Exception as exc:

            QMessageBox.critical(
                self,
                "Error",
                f"Could not add key:\n{exc}"
            )


    # ========================================================
    # DELETE KEY
    # ========================================================

    def delete_key(self):

        nickname = (
            self.parent_window
            .key_dropdown
            .currentText()
            .strip()
        )

        if not nickname:

            QMessageBox.warning(
                self,
                "Delete Key",
                "No key is selected."
            )

            return

        answer = QMessageBox.question(
            self,
            "Confirm Delete",
            f"Delete the key '{nickname}'?",
            QMessageBox.StandardButton.Yes
            | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No
        )

        if (
            answer
            != QMessageBox.StandardButton.Yes
        ):

            return

        try:

            wb = load_workbook(
                KEY_FILE
            )

            ws = wb.active

            target_row = None

            # ------------------------------------------------
            # Find Nickname column
            # ------------------------------------------------

            headers = [
                str(
                    ws.cell(
                        1,
                        c
                    ).value or ""
                )
                .strip()
                .lower()

                for c in range(
                    1,
                    ws.max_column + 1
                )
            ]

            try:

                nickname_col = (
                    headers.index(
                        "nickname"
                    )
                    + 1
                )

            except ValueError:

                # Backward compatibility
                nickname_col = 2

            # ------------------------------------------------
            # Find matching nickname
            # ------------------------------------------------

            for row in range(
                2,
                ws.max_row + 1
            ):

                value = ws.cell(
                    row,
                    nickname_col
                ).value

                if (
                    str(
                        value or ""
                    )
                    .strip()
                    == nickname
                ):

                    target_row = row

                    break

            if target_row is None:

                wb.close()

                QMessageBox.warning(
                    self,
                    "Not Found",
                    f"Could not find '{nickname}' in keys.xlsx."
                )

                return

            # ------------------------------------------------
            # Delete row
            # ------------------------------------------------

            ws.delete_rows(
                target_row,
                1
            )

            wb.save(
                KEY_FILE
            )

            wb.close()

            self.parent_window.load_data()

            QMessageBox.information(
                self,
                "Deleted",
                "API key deleted successfully."
            )

        except PermissionError:

            QMessageBox.critical(
                self,
                "Excel File Locked",
                "Close keys.xlsx in Excel and try again."
            )

        except Exception as exc:

            QMessageBox.critical(
                self,
                "Error",
                f"Could not delete key:\n{exc}"
            )


    # ========================================================
    # PROVIDER SETTINGS
    # ========================================================

    def provider_settings(self):

        dialog = ProviderSettingsDialog(
            self.parent_window,
            self.parent_window.providers
        )

        dialog.exec()


    # ========================================================
    # API TESTER
    # ========================================================

    def open_api_tester(self):

        import subprocess

        tester_script = Path(__file__).with_name("api_tester.py")

        if tester_script.exists():

            subprocess.Popen(
                [
                    sys.executable,
                    str(tester_script)
                ]
            )

        else:

            QMessageBox.warning(
                self,
                "API Tester Not Found",
                f"Could not find:\n{tester_script}"
            )


# ============================================================
# MAIN APPLICATION
# ============================================================

class KeySwitcher(QMainWindow):

    def __init__(self):

        super().__init__()

        self.keys = {}

        self.key_providers = {}

        self.providers = {
            "providers": {}
        }

        self.setWindowTitle(
            "Key Switcher"
        )

        self.resize(
            620,
            380
        )

        self.setup_ui()

        self.load_data()


    # ========================================================
    # UI
    # ========================================================

    def setup_ui(self):

        central = QWidget()

        self.setCentralWidget(
            central
        )

        main_layout = QVBoxLayout(
            central
        )
        main_layout.setContentsMargins(
            22,
            20,
            22,
            20
        )
        main_layout.setSpacing(
            14
        )

        # ----------------------------------------------------
        # TITLE & SUBTITLE
        # ----------------------------------------------------

        header_layout = QVBoxLayout()
        header_layout.setSpacing(2)

        title = QLabel(
            "Key Switcher"
        )
        title.setAlignment(
            Qt.AlignmentFlag.AlignCenter
        )
        title.setStyleSheet(
            "font-size: 24px; font-weight: bold; color: #f8fafc;"
        )

        subtitle = QLabel(
            "Fast Claude Code API Key & Model Switcher"
        )
        subtitle.setAlignment(
            Qt.AlignmentFlag.AlignCenter
        )
        subtitle.setStyleSheet(
            "font-size: 12px; color: #94a3b8; margin-bottom: 4px;"
        )

        header_layout.addWidget(title)
        header_layout.addWidget(subtitle)
        main_layout.addLayout(header_layout)

        # ----------------------------------------------------
        # CARD CONTAINER
        # ----------------------------------------------------

        card = QFrame()
        card.setObjectName("cardFrame")

        card_layout = QVBoxLayout(card)
        card_layout.setSpacing(12)

        form = QFormLayout()
        form.setSpacing(12)

        # Nickname
        self.key_dropdown = QComboBox()
        self.key_dropdown.setView(QListView())
        self.key_dropdown.currentTextChanged.connect(
            self.on_nickname_changed
        )

        form.addRow(
            "🔑  Nickname:",
            self.key_dropdown
        )

        # Model
        self.model_dropdown = QComboBox()
        self.model_dropdown.setView(QListView())

        form.addRow(
            "🤖  Model:",
            self.model_dropdown
        )

        card_layout.addLayout(form)

        # Live info preview
        self.preview_lbl = QLabel(
            "Provider: —  •  Base URL: —"
        )
        self.preview_lbl.setStyleSheet(
            "color: #94a3b8; font-size: 11px; padding: 4px 2px;"
        )
        card_layout.addWidget(self.preview_lbl)

        main_layout.addWidget(card)

        # ----------------------------------------------------
        # BUTTONS
        # ----------------------------------------------------

        buttons = QHBoxLayout()
        buttons.setSpacing(10)

        menu_btn = QPushButton(
            "⚙  Manage Keys & Providers"
        )
        menu_btn.setObjectName("secondaryBtn")
        menu_btn.setMinimumHeight(42)
        menu_btn.setToolTip(
            "Add/delete keys, configure providers, or launch API tester"
        )
        menu_btn.clicked.connect(
            self.open_menu
        )

        apply_btn = QPushButton(
            "✓  APPLY TO CLAUDE CODE"
        )
        apply_btn.setObjectName("applyBtn")
        apply_btn.setMinimumHeight(42)
        apply_btn.clicked.connect(
            self.apply_settings
        )

        buttons.addWidget(menu_btn, stretch=1)
        buttons.addWidget(apply_btn, stretch=2)

        main_layout.addLayout(
            buttons
        )

        # ----------------------------------------------------
        # STATUS
        # ----------------------------------------------------

        self.status = QLabel(
            "Ready • Select a key and model to apply"
        )
        self.status.setAlignment(
            Qt.AlignmentFlag.AlignCenter
        )
        self.status.setStyleSheet(
            "background-color: #1a1c2b; border: 1px solid #2d3148; border-radius: 6px; padding: 6px 12px; color: #94a3b8; font-size: 11px;"
        )

        main_layout.addWidget(
            self.status
        )


    # ========================================================
    # LOAD EVERYTHING
    # ========================================================

    def load_data(self):

        try:

            self.providers = load_providers()

            self.load_keys()

        except Exception as exc:

            QMessageBox.critical(
                self,
                "Startup Error",
                str(exc)
            )


    # ========================================================
    # LOAD EXCEL KEYS
    # ========================================================

    def load_keys(self):

        current = (
            self.key_dropdown
            .currentText()
        )

        self.keys.clear()

        self.key_providers.clear()

        self.key_dropdown.blockSignals(
            True
        )

        self.key_dropdown.clear()

        if not KEY_FILE.exists():

            self.key_dropdown.blockSignals(
                False
            )

            return

        try:

            wb = load_workbook(
                KEY_FILE,
                read_only=True,
                data_only=True
            )

            ws = wb.active

            # ------------------------------------------------
            # Read headers
            # ------------------------------------------------

            headers = [
                str(
                    ws.cell(
                        1,
                        c
                    ).value or ""
                )
                .strip()
                .lower()

                for c in range(
                    1,
                    ws.max_column + 1
                )
            ]

            # ------------------------------------------------
            # Find columns
            # ------------------------------------------------

            try:

                provider_col = (
                    headers.index(
                        "provider"
                    )
                    + 1
                )

            except ValueError:

                provider_col = None

            try:

                nickname_col = (
                    headers.index(
                        "nickname"
                    )
                    + 1
                )

            except ValueError:

                nickname_col = 2

            try:

                key_col = (
                    headers.index(
                        "api key"
                    )
                    + 1
                )

            except ValueError:

                # Old format:
                # Nickname | Key

                try:

                    key_col = (
                        headers.index(
                            "key"
                        )
                        + 1
                    )

                except ValueError:

                    key_col = 3

            # ------------------------------------------------
            # Read rows
            # ------------------------------------------------

            for row in ws.iter_rows(
                min_row=2,
                values_only=True
            ):

                # Provider
                provider = ""

                if provider_col:

                    if (
                        len(row)
                        >= provider_col
                        and row[
                            provider_col - 1
                        ]
                    ):

                        provider = str(
                            row[
                                provider_col - 1
                            ]
                        ).strip()

                # Nickname
                if (
                    len(row)
                    < nickname_col
                    or not row[
                        nickname_col - 1
                    ]
                ):

                    continue

                nickname = str(
                    row[
                        nickname_col - 1
                    ]
                ).strip()

                # API Key
                if (
                    len(row)
                    < key_col
                    or not row[
                        key_col - 1
                    ]
                ):

                    continue

                api_key = str(
                    row[
                        key_col - 1
                    ]
                ).strip()

                if nickname and api_key:

                    self.keys[
                        nickname
                    ] = api_key

                    self.key_providers[
                        nickname
                    ] = provider

            wb.close()

        except Exception as exc:

            self.key_dropdown.blockSignals(
                False
            )

            QMessageBox.critical(
                self,
                "Excel Error",
                f"Could not read keys.xlsx:\n{exc}"
            )

            return

        # ----------------------------------------------------
        # Populate Nickname dropdown
        # ----------------------------------------------------

        self.key_dropdown.addItems(
            self.keys.keys()
        )

        if (
            current
            and current in self.keys
        ):

            self.key_dropdown.setCurrentText(
                current
            )

        elif self.key_dropdown.count():

            self.key_dropdown.setCurrentIndex(
                0
            )

        self.key_dropdown.blockSignals(
            False
        )

        self.on_nickname_changed(
            self.key_dropdown.currentText()
        )


    # ========================================================
    # NICKNAME CHANGED
    #
    # This determines the provider from Excel.
    # Then it loads that provider's models from JSON.
    # ========================================================

    def on_nickname_changed(
        self,
        nickname
    ):

        self.model_dropdown.clear()

        provider = (
            self.key_providers
            .get(
                nickname,
                ""
            )
            .strip()
        )

        # ----------------------------------------------------
        # Provider not found
        # ----------------------------------------------------

        if not provider:

            if hasattr(self, "preview_lbl"):
                self.preview_lbl.setText("Provider: Not found  •  Base URL: —")

            self.status.setText(
                "Provider not found for selected nickname"
            )

            return

        # ----------------------------------------------------
        # Find provider in providers.json
        # ----------------------------------------------------

        provider_data = (
            self.providers[
                "providers"
            ].get(
                provider
            )
        )

        if provider_data is None:

            if hasattr(self, "preview_lbl"):
                self.preview_lbl.setText(f"Provider '{provider}' not configured in providers.json")

            self.status.setText(
                f"Provider '{provider}' not found in providers.json"
            )

            return

        # ----------------------------------------------------
        # Get models
        # ----------------------------------------------------

        models = provider_data.get(
            "models",
            []
        )

        self.model_dropdown.addItems(
            models
        )

        base_url = provider_data.get("base_url", "None")
        if hasattr(self, "preview_lbl"):
            self.preview_lbl.setText(
                f"Provider: <b style='color:#f8fafc;'>{provider}</b> &nbsp;•&nbsp; "
                f"Base URL: <code style='color:#a5b4fc;'>{base_url}</code>"
            )

        # ----------------------------------------------------
        # Show status
        # ----------------------------------------------------

        self.status.setText(
            f"Ready: {provider} • {len(models)} model(s) available"
        )


    # ========================================================
    # OPEN MENU
    # ========================================================

    def open_menu(self):

        dialog = KeyManagerDialog(
            self
        )

        dialog.exec()


    # ========================================================
    # APPLY SETTINGS
    # ========================================================

    def apply_settings(self):

        # ----------------------------------------------------
        # Selected nickname
        # ----------------------------------------------------

        nickname = (
            self.key_dropdown
            .currentText()
            .strip()
        )

        # ----------------------------------------------------
        # Selected model
        # ----------------------------------------------------

        model = (
            self.model_dropdown
            .currentText()
            .strip()
        )

        # ----------------------------------------------------
        # Validate nickname
        # ----------------------------------------------------

        if not nickname:

            QMessageBox.warning(
                self,
                "Missing Key",
                "Select a nickname."
            )

            return

        # ----------------------------------------------------
        # Validate model
        # ----------------------------------------------------

        if not model:

            QMessageBox.warning(
                self,
                "Missing Model",
                "Select a model."
            )

            return

        # ----------------------------------------------------
        # Get API key
        # ----------------------------------------------------

        api_key = self.keys.get(
            nickname
        )

        if not api_key:

            QMessageBox.warning(
                self,
                "Missing Key",
                "The selected API key was not found."
            )

            return

        # ----------------------------------------------------
        # Get provider from Excel
        # ----------------------------------------------------

        provider = (
            self.key_providers
            .get(
                nickname,
                ""
            )
            .strip()
        )

        if not provider:

            QMessageBox.warning(
                self,
                "Missing Provider",
                "No provider is associated with this nickname."
            )

            return

        # ----------------------------------------------------
        # Get provider configuration from JSON
        # ----------------------------------------------------

        provider_data = (
            self.providers[
                "providers"
            ].get(
                provider
            )
        )

        if provider_data is None:

            QMessageBox.critical(
                self,
                "Provider Not Found",
                f"Provider '{provider}' was not found "
                "in providers.json."
            )

            return

        # ----------------------------------------------------
        # Get BASE URL from JSON
        # ----------------------------------------------------

        base_url = (
            provider_data
            .get(
                "base_url",
                ""
            )
            .strip()
        )

        if not base_url:

            QMessageBox.critical(
                self,
                "Missing Base URL",
                f"No base_url is configured for "
                f"provider '{provider}'."
            )

            return

        # ----------------------------------------------------
        # Check Claude Code settings
        # ----------------------------------------------------

        if not SETTINGS_FILE.exists():

            QMessageBox.critical(
                self,
                "Settings Not Found",
                "Claude Code settings file was not found:\n"
                f"{SETTINGS_FILE}"
            )

            return

        # ----------------------------------------------------
        # UPDATE SETTINGS.JSON
        # ----------------------------------------------------

        try:

            with SETTINGS_FILE.open(
                "r",
                encoding="utf-8"
            ) as file:

                settings = json.load(
                    file
                )

            if not isinstance(
                settings,
                dict
            ):

                raise ValueError(
                    "settings.json must contain a JSON object."
                )

            # ------------------------------------------------
            # Get env
            # ------------------------------------------------

            env = settings.setdefault(
                "env",
                {}
            )

            if not isinstance(
                env,
                dict
            ):

                raise ValueError(
                    "'env' in settings.json must be an object."
                )

            # ------------------------------------------------
            # API KEY
            # ------------------------------------------------

            env[
                "ANTHROPIC_AUTH_TOKEN"
            ] = api_key

            # ------------------------------------------------
            # BASE URL
            #
            # THIS IS THE IMPORTANT FIX
            # ------------------------------------------------

            env[
                "ANTHROPIC_BASE_URL"
            ] = base_url

            # ------------------------------------------------
            # MODELS
            # ------------------------------------------------

            for setting in MODEL_SETTINGS:

                env[
                    setting
                ] = model

            # ------------------------------------------------
            # SAVE SETTINGS
            # ------------------------------------------------

            with SETTINGS_FILE.open(
                "w",
                encoding="utf-8"
            ) as file:

                json.dump(
                    settings,
                    file,
                    indent=2,
                    ensure_ascii=False
                )

                file.write(
                    "\n"
                )

            save_switch_history(
                nickname,
                model
            )

            # ------------------------------------------------
            # SUCCESS
            # ------------------------------------------------

            self.status.setText(
                f"Applied: {provider} • {nickname} • {model}"
            )

            QMessageBox.information(
                self,
                "Applied",
                "Claude Code settings updated.\n\n"
                f"Provider: {provider}\n"
                f"Nickname: {nickname}\n"
                f"Model: {model}\n"
                f"Base URL: {base_url}"
            )

        # ----------------------------------------------------
        # ERRORS
        # ----------------------------------------------------

        except json.JSONDecodeError:

            QMessageBox.critical(
                self,
                "Invalid JSON",
                "settings.json contains invalid JSON."
            )

        except PermissionError:

            QMessageBox.critical(
                self,
                "Permission Error",
                "Could not write settings.json."
            )

        except Exception as exc:

            QMessageBox.critical(
                self,
                "Error",
                f"Could not update Claude Code settings:\n{exc}"
            )


# ============================================================
# MAIN
# ============================================================

def main():

    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    app.setStyleSheet(APP_STYLESHEET)

    window = KeySwitcher()

    window.show()

    sys.exit(
        app.exec()
    )


if __name__ == "__main__":

    main()
### it have multiple provder support


import sys
import json
from pathlib import Path

from openpyxl import load_workbook, Workbook

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QApplication,
    QComboBox,
    QDialog,
    QFormLayout,
    QHBoxLayout,
    QInputDialog,
    QLabel,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QVBoxLayout,
    QWidget,
    QListWidget,
    QLineEdit,
)

SETTINGS_FILE = Path(r"C:\Users\Nishant\.claude\settings.json")
KEY_FILE = Path(r"C:\My Space\Data\keys.xlsx")
PROVIDERS_FILE = Path(__file__).with_name("providers.json")

MODEL_SETTINGS = [
    "ANTHROPIC_MODEL",
    "ANTHROPIC_SMALL_FAST_MODEL",
    "ANTHROPIC_DEFAULT_SONNET_MODEL",
    "ANTHROPIC_DEFAULT_OPUS_MODEL",
    "ANTHROPIC_DEFAULT_HAIKU_MODEL",
]

DEFAULT_PROVIDERS = {
    "providers": {
        "Ollama": {
            "base_url": "http://localhost:11434",
            "models": [
                "gemma4:31b",
                "gpt-oss:120b",
                "gpt-oss:20b",
                "nemotron-3-nano:30b",
                "nemotron-3-super",
                "nemotron-3-ultra",
            ],
        },
        "Hugging Face": {
            "base_url": "https://router.huggingface.co/v1",
            "models": [],
        },
    }
}


def ensure_provider_file():
    if not PROVIDERS_FILE.exists():
        PROVIDERS_FILE.write_text(
            json.dumps(DEFAULT_PROVIDERS, indent=2, ensure_ascii=False),
            encoding="utf-8",
        )


def load_providers():
    ensure_provider_file()

    try:
        data = json.loads(
            PROVIDERS_FILE.read_text(encoding="utf-8")
        )

        if not isinstance(data.get("providers"), dict):
            raise ValueError("Invalid providers.json format.")

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
        encoding="utf-8",
    )


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

        wb.save(KEY_FILE)
        wb.close()


class ProviderSettingsDialog(QDialog):

    def __init__(self, parent, providers):
        super().__init__(parent)

        self.parent_window = parent
        self.providers = providers

        self.setWindowTitle("Provider Settings")
        self.resize(650, 450)

        layout = QVBoxLayout(self)

        form = QFormLayout()

        self.provider_dropdown = QComboBox()

        self.provider_dropdown.addItems(
            self.providers["providers"].keys()
        )

        self.provider_dropdown.currentTextChanged.connect(
            self.load_provider
        )

        self.url_edit = QLineEdit()

        self.url_edit.setPlaceholderText(
            "Base URL"
        )

        form.addRow(
            "Provider:",
            self.provider_dropdown
        )

        form.addRow(
            "Base URL:",
            self.url_edit
        )

        layout.addLayout(form)

        layout.addWidget(
            QLabel("Models:")
        )

        self.model_list = QListWidget()

        layout.addWidget(
            self.model_list
        )

        buttons = QHBoxLayout()

        add_btn = QPushButton(
            "Add Model"
        )

        delete_btn = QPushButton(
            "Delete Model"
        )

        save_btn = QPushButton(
            "SAVE"
        )

        close_btn = QPushButton(
            "CLOSE"
        )

        add_btn.clicked.connect(
            self.add_model
        )

        delete_btn.clicked.connect(
            self.delete_model
        )

        save_btn.clicked.connect(
            self.save
        )

        close_btn.clicked.connect(
            self.reject
        )

        buttons.addWidget(add_btn)
        buttons.addWidget(delete_btn)
        buttons.addStretch()
        buttons.addWidget(save_btn)
        buttons.addWidget(close_btn)

        layout.addLayout(buttons)

        if self.provider_dropdown.count():
            self.load_provider(
                self.provider_dropdown.currentText()
            )

    def load_provider(self, provider):

        data = self.providers[
            "providers"
        ].get(provider, {})

        self.url_edit.setText(
            data.get("base_url", "")
        )

        self.model_list.clear()

        self.model_list.addItems(
            data.get("models", [])
        )

    def add_model(self):

        model, ok = QInputDialog.getText(
            self,
            "Add Model",
            "Model name:"
        )

        model = model.strip()

        if not ok or not model:
            return

        existing = [
            self.model_list.item(i).text()
            for i in range(
                self.model_list.count()
            )
        ]

        if model in existing:

            QMessageBox.warning(
                self,
                "Duplicate",
                "This model already exists."
            )

            return

        self.model_list.addItem(model)

    def delete_model(self):

        item = self.model_list.currentItem()

        if not item:

            QMessageBox.warning(
                self,
                "Delete Model",
                "Select a model first."
            )

            return

        self.model_list.takeItem(
            self.model_list.row(item)
        )

    def save(self):

        provider = (
            self.provider_dropdown
            .currentText()
        )

        models = [
            self.model_list.item(i).text().strip()
            for i in range(
                self.model_list.count()
            )
            if self.model_list.item(i).text().strip()
        ]

        self.providers[
            "providers"
        ][provider] = {

            "base_url":
                self.url_edit.text().strip(),

            "models":
                models
        }

        try:

            save_providers(
                self.providers
            )

            self.parent_window.load_data()

            QMessageBox.information(
                self,
                "Saved",
                "Provider settings updated."
            )

            self.accept()

        except Exception as exc:

            QMessageBox.critical(
                self,
                "Error",
                f"Could not save providers.json:\n{exc}"
            )


class KeyManagerDialog(QDialog):

    def __init__(self, parent):

        super().__init__(parent)

        self.parent_window = parent

        self.setWindowTitle(
            "Key Manager"
        )

        self.resize(
            350,
            200
        )

        layout = QVBoxLayout(self)

        add_btn = QPushButton(
            "Add Key"
        )

        delete_btn = QPushButton(
            "Delete Selected Key"
        )

        provider_btn = QPushButton(
            "Provider Settings"
        )

        close_btn = QPushButton(
            "Close"
        )

        add_btn.clicked.connect(
            self.add_key
        )

        delete_btn.clicked.connect(
            self.delete_key
        )

        provider_btn.clicked.connect(
            self.provider_settings
        )

        close_btn.clicked.connect(
            self.accept
        )

        layout.addWidget(add_btn)
        layout.addWidget(delete_btn)
        layout.addWidget(provider_btn)
        layout.addWidget(close_btn)

    def add_key(self):

        providers = list(
            self.parent_window
            .providers["providers"]
            .keys()
        )

        if not providers:

            QMessageBox.warning(
                self,
                "No Providers",
                "Add a provider first."
            )

            return

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

        try:

            ensure_key_file()

            wb = load_workbook(
                KEY_FILE
            )

            ws = wb.active

            if ws.max_row == 1 and all(
                ws.cell(1, c).value is None
                for c in range(1, 4)
            ):

                ws.append([
                    "Provider",
                    "Nickname",
                    "API Key"
                ])

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
                "Close key.xlsx in Excel and try again."
            )

        except Exception as exc:

            QMessageBox.critical(
                self,
                "Error",
                f"Could not add key:\n{exc}"
            )

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

        if answer != QMessageBox.StandardButton.Yes:
            return

        try:

            ensure_key_file()

            wb = load_workbook(
                KEY_FILE
            )

            ws = wb.active

            target_row = None

            headers = [
                str(
                    ws.cell(1, c).value or ""
                ).strip().lower()
                for c in range(
                    1,
                    min(ws.max_column, 3) + 1
                )
            ]

            if headers[:3] == [
                "provider",
                "nickname",
                "api key"
            ]:

                nickname_col = 2

            elif headers[:2] == [
                "nickname",
                "key"
            ]:

                nickname_col = 1

            else:

                nickname_col = (
                    2
                    if ws.max_column >= 3
                    else 1
                )

            for row in range(
                2,
                ws.max_row + 1
            ):

                value = ws.cell(
                    row,
                    nickname_col
                ).value

                if (
                    str(value or "").strip()
                    == nickname
                ):

                    target_row = row
                    break

            if target_row is None:

                wb.close()

                QMessageBox.warning(
                    self,
                    "Not Found",
                    f"Could not find '{nickname}' in key.xlsx."
                )

                return

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
                "Close key.xlsx in Excel and try again."
            )

        except Exception as exc:

            QMessageBox.critical(
                self,
                "Error",
                f"Could not delete key:\n{exc}"
            )

    def provider_settings(self):

        dialog = ProviderSettingsDialog(
            self.parent_window,
            self.parent_window.providers
        )

        dialog.exec()


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
            560,
            260
        )

        self.setup_ui()

        self.load_data()

    def setup_ui(self):

        central = QWidget()

        self.setCentralWidget(
            central
        )

        main_layout = QVBoxLayout(
            central
        )

        title = QLabel(
            "Key Switcher"
        )

        title.setAlignment(
            Qt.AlignmentFlag.AlignCenter
        )

        title.setStyleSheet(
            "font-size: 22px; font-weight: bold;"
        )

        main_layout.addWidget(
            title
        )

        form = QFormLayout()

        self.key_dropdown = QComboBox()

        self.key_dropdown.currentTextChanged.connect(
            self.on_nickname_changed
        )

        self.model_dropdown = QComboBox()

        form.addRow(
            "Nickname:",
            self.key_dropdown
        )

        form.addRow(
            "Model:",
            self.model_dropdown
        )

        main_layout.addLayout(
            form
        )

        buttons = QHBoxLayout()

        menu_btn = QPushButton(
            "⋮"
        )

        menu_btn.setFixedWidth(
            45
        )

        menu_btn.setToolTip(
            "Key and provider management"
        )

        menu_btn.clicked.connect(
            self.open_menu
        )

        apply_btn = QPushButton(
            "APPLY"
        )

        apply_btn.setMinimumHeight(
            38
        )

        apply_btn.clicked.connect(
            self.apply_settings
        )

        buttons.addStretch()

        buttons.addWidget(
            menu_btn
        )

        buttons.addWidget(
            apply_btn
        )

        main_layout.addLayout(
            buttons
        )

        self.status = QLabel(
            "Ready"
        )

        self.status.setAlignment(
            Qt.AlignmentFlag.AlignCenter
        )

        main_layout.addWidget(
            self.status
        )

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

    def load_keys(self):

        current = (
            self.key_dropdown.currentText()
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

            headers = [
                str(
                    ws.cell(1, c).value or ""
                ).strip().lower()
                for c in range(
                    1,
                    min(ws.max_column, 3) + 1
                )
            ]

            if headers[:3] == [
                "provider",
                "nickname",
                "api key"
            ]:

                provider_col = 1
                nickname_col = 2
                key_col = 3

            elif headers[:2] == [
                "nickname",
                "key"
            ]:

                provider_col = None
                nickname_col = 1
                key_col = 2

            else:

                provider_col = 1
                nickname_col = 2
                key_col = 3

            for row in ws.iter_rows(
                min_row=2,
                values_only=True
            ):

                provider = ""

                if (
                    provider_col
                    and len(row) >= provider_col
                    and row[provider_col - 1]
                ):

                    provider = str(
                        row[provider_col - 1]
                    ).strip()

                if (
                    len(row) < nickname_col
                    or not row[nickname_col - 1]
                ):

                    continue

                if (
                    len(row) < key_col
                    or not row[key_col - 1]
                ):

                    continue

                nickname = str(
                    row[nickname_col - 1]
                ).strip()

                api_key = str(
                    row[key_col - 1]
                ).strip()

                if nickname and api_key:

                    self.keys[nickname] = api_key

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
                f"Could not read key.xlsx:\n{exc}"
            )

            return

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

    def on_nickname_changed(
        self,
        nickname
    ):

        self.model_dropdown.clear()

        provider = self.key_providers.get(
            nickname,
            ""
        ).strip()

        if not provider:

            if len(
                self.providers["providers"]
            ) == 1:

                provider = next(
                    iter(
                        self.providers["providers"]
                    )
                )

        provider_data = (
            self.providers["providers"]
            .get(
                provider,
                {}
            )
        )

        models = provider_data.get(
            "models",
            []
        )

        self.model_dropdown.addItems(
            models
        )

        if provider:

            self.status.setText(
                f"{provider} • {len(models)} model(s)"
            )

        else:

            self.status.setText(
                "Select a nickname"
            )

    def open_menu(self):

        dialog = KeyManagerDialog(
            self
        )

        dialog.exec()

    def apply_settings(self):

        nickname = (
            self.key_dropdown
            .currentText()
            .strip()
        )

        model = (
            self.model_dropdown
            .currentText()
            .strip()
        )

        if not nickname:

            QMessageBox.warning(
                self,
                "Missing Key",
                "Select a nickname."
            )

            return

        if not model:

            QMessageBox.warning(
                self,
                "Missing Model",
                "Select a model."
            )

            return

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

        if not SETTINGS_FILE.exists():

            QMessageBox.critical(
                self,
                "Settings Not Found",
                "Claude Code settings file was not found:\n"
                f"{SETTINGS_FILE}"
            )

            return

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

            env[
                "ANTHROPIC_AUTH_TOKEN"
            ] = api_key

            for setting in MODEL_SETTINGS:

                env[setting] = model

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

                file.write("\n")

            provider = (
                self.key_providers
                .get(
                    nickname,
                    "Unknown provider"
                )
            )

            self.status.setText(
                f"Applied: {nickname} • {model}"
            )

            QMessageBox.information(
                self,
                "Applied",
                "Claude Code settings updated.\n\n"
                f"Provider: {provider}\n"
                f"Nickname: {nickname}\n"
                f"Model: {model}"
            )

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


def main():

    app = QApplication(
        sys.argv
    )

    window = KeySwitcher()

    window.show()

    sys.exit(
        app.exec()
    )


if __name__ == "__main__":
    main()
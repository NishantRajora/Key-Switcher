# Key Switcher

A lightweight Windows GUI application for quickly switching **Claude Code API keys, providers, base URLs, and models**.

Key Switcher uses an Excel file to store API keys and a `providers.json` file to manage provider-specific configuration. When you click **APPLY**, the selected configuration is written directly to the Claude Code `settings.json` file.

---

## Features

* 🔑 Store and manage multiple API keys
* 🏷️ Give each API key a custom nickname
* 🔌 Support multiple providers
* 🤖 Provider-specific model lists
* 🌐 Provider-specific base URLs
* ⚡ Quickly switch between API keys and models
* 📝 Automatically update Claude Code `settings.json`
* 📊 Store API keys in an Excel file
* ⚙️ Edit providers and models from the application
* 🗑️ Delete saved API keys
* 🚫 No backup files are created
* 🎯 Simple main interface with only the required controls

---

## How It Works

Key Switcher uses three main configuration sources:

```text
                 ┌──────────────────┐
                 │    keys.xlsx     │
                 │                  │
                 │ Provider         │
                 │ Nickname         │
                 │ API Key          │
                 └────────┬─────────┘
                          │
                          ▼
                   Selected Nickname
                          │
                          ▼
                    Find Provider
                          │
                          ▼
                 ┌──────────────────┐
                 │  providers.json  │
                 │                  │
                 │ Base URL         │
                 │ Models           │
                 └────────┬─────────┘
                          │
                          ▼
                  Selected Model
                          │
                          ▼
                 ┌──────────────────┐
                 │  settings.json   │
                 │                  │
                 │ Auth Token        │
                 │ Base URL         │
                 │ Model Settings   │
                 └──────────────────┘
```

The provider is **not selected from the main interface**.

Instead:

1. You select a **Nickname**.
2. Key Switcher finds the provider associated with that nickname in `keys.xlsx`.
3. It loads that provider's models from `providers.json`.
4. You select a model.
5. Clicking **APPLY** updates Claude Code's `settings.json`.

---

# Project Structure

A typical project structure is:

```text
Key Switcher/
│
├── Switcher_v2.py
├── providers.json
└── README.md
```

The Excel key file is stored separately:

```text
C:\My Space\Data\keys.xlsx
```

Claude Code's settings file is:

```text
C:\Users\Nishant\.claude\settings.json
```

---

# Requirements

## Python

Python 3.10+ is recommended.

Check your Python installation:

```powershell
python --version
```

---

## Required Python Packages

Install the required packages with:

```powershell
pip install PySide6 openpyxl
```

The application uses:

* **PySide6** — GUI
* **openpyxl** — Excel file management
* **json** — provider/settings configuration
* **pathlib** — file path management

`json` and `pathlib` are included with Python.

---

# Installation

Clone or download the project.

Then open a terminal in the project directory:

```powershell
cd "C:\path\to\Key Switcher"
```

Install dependencies:

```powershell
pip install PySide6 openpyxl
```

Run the application:

```powershell
python Switcher_v2.py
```

---

# Configuration Files

## 1. keys.xlsx

API keys are stored in:

```text
C:\My Space\Data\keys.xlsx
```

The Excel file contains only three columns:

| Provider     | Nickname    | API Key      |
| ------------ | ----------- | ------------ |
| Ollama       | Ollama Main | your-api-key |
| Hugging Face | HF Key      | your-api-key |

### Important

The application does **not** store the model or base URL in Excel.

Those values come from `providers.json`.

The provider is stored with each key so that Key Switcher knows which provider configuration to use.

---

# 2. providers.json

`providers.json` is located in the same directory as the Python application.

Example:

```json
{
  "providers": {
    "Ollama": {
      "base_url": "https://ollama.com",
      "models": [
        "gemma4:31b",
        "gpt-oss:120b",
        "gpt-oss:20b",
        "nemotron-3-nano:30b",
        "nemotron-3-super",
        "nemotron-3-ultra"
      ]
    },
    "Hugging Face": {
      "base_url": "https://router.huggingface.co/v1",
      "models": [
        "Qwen/Qwen3.8-27B"
      ]
    }
  }
}
```

Each provider contains:

```text
Provider
 ├── base_url
 └── models
```

### Base URL

The `base_url` is provider-wide.

For example:

```json
"Ollama": {
  "base_url": "https://ollama.com"
}
```

All keys belonging to Ollama use that base URL.

Similarly:

```json
"Hugging Face": {
  "base_url": "https://router.huggingface.co/v1"
}
```

All Hugging Face keys use that base URL.

---

# Adding a Provider

Open the application and click the:

```text
⋮
```

menu.

Select:

```text
Provider Settings
```

From there you can edit the provider configuration.

You can:

* Change the Base URL
* Add models
* Delete models

The changes are saved directly to:

```text
providers.json
```

---

# Adding an API Key

Click:

```text
⋮
```

Then select:

```text
Add Key
```

The application asks for:

1. Provider
2. Nickname
3. API Key

Example:

```text
Provider: Ollama
Nickname: My Ollama
API Key: ********
```

The resulting Excel row is:

```text
Ollama | My Ollama | ********
```

---

# Switching API Keys

The main interface contains:

```text
Nickname: [ My Ollama ▼ ]

Model:    [ gemma4:31b ▼ ]

                         ⋮  APPLY
```

When you select a nickname, the application automatically determines the provider.

For example:

```text
Nickname
    │
    ▼
My Ollama
    │
    ▼
Provider = Ollama
    │
    ▼
providers.json
    │
    ├── Base URL
    │
    └── Models
```

The model dropdown is then populated with the models configured for that provider.

---

# Applying Settings

After selecting:

```text
Nickname
```

and:

```text
Model
```

click:

```text
APPLY
```

The application updates:

```text
C:\Users\Nishant\.claude\settings.json
```

The following environment variables are updated:

```text
ANTHROPIC_AUTH_TOKEN
ANTHROPIC_BASE_URL
ANTHROPIC_MODEL
ANTHROPIC_SMALL_FAST_MODEL
ANTHROPIC_DEFAULT_SONNET_MODEL
ANTHROPIC_DEFAULT_OPUS_MODEL
ANTHROPIC_DEFAULT_HAIKU_MODEL
```

---

# Example

Suppose `keys.xlsx` contains:

```text
Provider       Nickname       API Key
Ollama         Work           abc123
Hugging Face   Testing        xyz789
```

And `providers.json` contains:

```json
{
  "providers": {
    "Ollama": {
      "base_url": "https://ollama.com",
      "models": [
        "gemma4:31b",
        "gpt-oss:120b"
      ]
    },
    "Hugging Face": {
      "base_url": "https://router.huggingface.co/v1",
      "models": [
        "Qwen/Qwen3.8-27B"
      ]
    }
  }
}
```

If you select:

```text
Nickname: Work
Model: gpt-oss:120b
```

Key Switcher determines:

```text
Provider = Ollama
Base URL = https://ollama.com
API Key = abc123
Model = gpt-oss:120b
```

Then it writes those values to Claude Code.

---

# settings.json

The application modifies the existing:

```text
C:\Users\Nishant\.claude\settings.json
```

It preserves the other existing settings and updates the `env` section.

For example:

```json
{
  "env": {
    "ANTHROPIC_AUTH_TOKEN": "your-api-key",
    "ANTHROPIC_BASE_URL": "https://ollama.com",
    "ANTHROPIC_MODEL": "gpt-oss:120b",
    "ANTHROPIC_SMALL_FAST_MODEL": "gpt-oss:120b",
    "ANTHROPIC_DEFAULT_SONNET_MODEL": "gpt-oss:120b",
    "ANTHROPIC_DEFAULT_OPUS_MODEL": "gpt-oss:120b",
    "ANTHROPIC_DEFAULT_HAIKU_MODEL": "gpt-oss:120b"
  }
}
```

Other settings already present in `settings.json` are preserved.

---

# Key Management

The **Key Manager** provides:

### Add Key

Adds a new provider, nickname, and API key to `keys.xlsx`.

### Delete Selected Key

Deletes the currently selected nickname from `keys.xlsx`.

### Provider Settings

Opens the provider configuration editor.

---

# Security

API keys are stored locally in:

```text
C:\My Space\Data\keys.xlsx
```

They are not intentionally uploaded by Key Switcher to a remote server.

However, anyone who can access the Excel file may be able to read the API keys.

### Recommended

Protect the Excel file and your Windows account appropriately.

Do not commit `keys.xlsx` containing real API keys to Git.

Add it to `.gitignore`:

```gitignore
keys.xlsx
*.xlsx
```

You should also avoid committing real API keys anywhere in the repository.

---

# No Backups

Key Switcher intentionally does **not** create backup copies of:

```text
settings.json
```

or:

```text
keys.xlsx
```

When changes are applied, the existing file is updated directly.

Make sure your important configuration is recoverable before making changes.

---

# Error Handling

The application handles several common errors.

## Settings file not found

If this file does not exist:

```text
C:\Users\Nishant\.claude\settings.json
```

the application will show an error.

---

## Excel file locked

If `keys.xlsx` is open in Microsoft Excel while trying to add or delete a key, close the Excel file and try again.

---

## Invalid JSON

If `settings.json` or `providers.json` contains invalid JSON, the application reports the problem instead of applying the configuration.

---

## Provider Not Found

If a provider stored in `keys.xlsx` does not exist in `providers.json`, the model list cannot be loaded and the application reports the missing provider.

---

## Missing Base URL

If the selected provider has no `base_url`, the application will not apply the configuration.

---

# Data Flow

The complete configuration flow is:

```text
                    ┌──────────────┐
                    │   keys.xlsx  │
                    └──────┬───────┘
                           │
                     Nickname
                           │
                           ▼
                  ┌─────────────────┐
                  │ Find Provider   │
                  └────────┬────────┘
                           │
                           ▼
                  ┌─────────────────┐
                  │ providers.json  │
                  └───────┬─────────┘
                          │
                 ┌────────┴────────┐
                 │                 │
              Base URL           Models
                 │                 │
                 │                 ▼
                 │          Model Dropdown
                 │                 │
                 └────────┬────────┘
                          │
                       APPLY
                          │
                          ▼
                ┌───────────────────┐
                │   settings.json   │
                └───────────────────┘
```

---

# Main Interface

The main interface intentionally contains only:

```text
Nickname
Model

⋮   APPLY
```

The provider is hidden from the main interface because it is automatically determined from the selected nickname.

---

# Project Design

The application separates key storage from provider configuration.

### Excel

Responsible for:

```text
Provider
Nickname
API Key
```

### providers.json

Responsible for:

```text
Provider
Base URL
Models
```

### settings.json

Responsible for:

```text
Claude Code environment configuration
```

This makes it possible to use multiple keys for the same provider without duplicating the provider's model and base URL configuration.

---

# Example Multi-Key Setup

You can have multiple keys for the same provider:

```text
Provider    Nickname       API Key
Ollama      Personal       key-1
Ollama      Work           key-2
Ollama      Testing        key-3
```

All three automatically use:

```json
"Ollama": {
  "base_url": "https://ollama.com",
  "models": [...]
}
```

You can also have keys for different providers:

```text
Provider       Nickname       API Key
Ollama         Personal       key-1
Ollama         Work           key-2
Hugging Face   HF Testing     key-3
```

Selecting `HF Testing` automatically switches the available model list to the models configured under `Hugging Face`.

---

# Troubleshooting

### Model dropdown is empty

Check that:

1. The nickname exists in `keys.xlsx`.
2. The provider is correctly written in the `Provider` column.
3. The provider exists in `providers.json`.
4. The provider has at least one model.

---

### APPLY does not work

Check:

1. `settings.json` exists.
2. `settings.json` contains valid JSON.
3. The selected nickname has an API key.
4. The nickname has a provider.
5. The provider exists in `providers.json`.
6. The provider has a valid `base_url`.
7. The selected model exists.

---

### Excel changes are not detected

Close `keys.xlsx` in Excel and restart Key Switcher if necessary.

---

# Running the Application

From the project directory:

```powershell
python Switcher_v2.py
```

If the Python command is not available, try:

```powershell
py Switcher_v2.py
```

---

# Future Improvements

Possible future improvements include:

* System tray support
* Global keyboard shortcuts
* Searchable model list
* Import/export provider configurations
* Encrypted API-key storage
* Windows executable build
* Automatic Claude Code restart/reload
* Connection testing
* Provider status indicators
* Dark/light theme support

---

# License

Add your preferred license here.

For example:

```text
MIT License
```

if you decide to release the project under the MIT License.

---

## Summary

Key Switcher provides a simple way to manage multiple Claude Code API configurations.

The architecture is:

```text
keys.xlsx
    │
    ├── Provider
    ├── Nickname
    └── API Key
         │
         ▼
providers.json
    │
    ├── Base URL
    └── Models
         │
         ▼
settings.json
    │
    ├── API Key
    ├── Base URL
    └── Model Configuration
```

Select a nickname, select a model, and press **APPLY** to switch the Claude Code configuration.

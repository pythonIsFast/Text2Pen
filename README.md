# Text2Pen

Text2Pen is a Windows tool that converts typed text into handwritten-style input directly inside Microsoft OneNote.

---

## 🖊️ What is Text2Pen?

Text2Pen lets you type any text and have it automatically "written" into OneNote as handwritten ink — no stylus required.
It is especially useful for students, note-takers, and anyone who wants the aesthetic of handwritten notes without the effort.

---

## ✨ Why is it cool?

- **Type, don't write** — produce handwritten-looking notes at typing speed
- **OneNote integration** — output goes directly into your open OneNote page
- **Table support** — paste tab-separated content and it is inserted as a table
- **AI sidebar** — chat assistant with image-upload support built right into the app *(backend temporarily unavailable — see Limitations)*
- **Automatic updates** — the updater keeps Text2Pen and itself up to date silently in the background, with SHA-256 hash verification for every download
- **Privacy-first** — completely offline by default; optional anonymous telemetry is opt-in only

---

## 🚀 Installation

1. Download the latest **Installer.exe** from the [Releases](https://github.com/pythonIsFast/Text2Pen/releases/latest) page.
2. Run `Installer.exe` and click **Install**.
   - `Text2Pen.exe` is placed in `%LOCALAPPDATA%\Text2Pen\`.
   - A Start-Menu shortcut is created automatically.
   - The updater (`Update.exe`) is registered to run on startup.
3. Launch **Text2Pen** from the Start Menu.

## ▶️ Usage

1. Open Microsoft OneNote and navigate to the page where you want to write.
2. Start **Text2Pen**.
3. Type (or paste) your text into the input field.
4. Click **Write** — Text2Pen will inject the text as handwritten ink into OneNote.

### Uninstallation

Run `Installer.exe` again and click **Uninstall**. All files and shortcuts are removed automatically.

---

## 📋 Features

| Feature | Details |
|---|---|
| Text → Handwriting | Converts typed text to handwritten ink in OneNote |
| Table input | Tab-separated text is inserted as a table |
| AI Chat Sidebar | Built-in assistant with model selection and image upload |
| Automatic Updates | Background updater with SHA-256 integrity verification |
| Optional Telemetry | Anonymous crash/error reporting (opt-in, disabled by default) |
| Windows-native | Deep integration via Win32 API |

---

## ⚠️ Limitations

- **Windows only** — Text2Pen uses the Win32 API and only runs on Windows.
- **Requires Microsoft OneNote** — the desktop version must be open and in focus when writing.
- **Backend / AI features temporarily unavailable** — some backend-dependent features (AI chat, AI text generation) are currently disabled due to legal reasons.

---

## ⚠️ Disclaimer

> **Use Text2Pen entirely at your own risk.**

- Text2Pen automates input into Microsoft OneNote, which **may violate Microsoft's Terms of Service**. Use it only for personal, non-commercial purposes.
- The tool may overwrite, modify, or delete OneNote content if used incorrectly.
- The developers are **not liable** for any loss of data, application crashes, or other unintended effects.
- Always **save your work** before using Text2Pen.
- Text2Pen may replace or update its executable files automatically. Only run it if you trust the source.

---

## 🔒 Privacy Policy

### 1. Overview

Text2Pen respects your privacy. The application works completely **offline by default**. No personal data is collected without your explicit consent.

### 2. Telemetry (Optional, Opt-In)

Text2Pen offers optional anonymous telemetry to help improve stability and reliability.

- Telemetry is **disabled by default** and only enabled if you explicitly opt in.
- You can change this setting at any time in the application settings.

### 3. What data is collected (if enabled)

If telemetry is enabled, the following anonymous technical data may be sent:

- Application error and crash reports
- Operating system type (e.g. Windows)
- Timestamp of the error
- Sanitized error messages (no usernames, no file contents)

### 4. What data is NOT collected

Text2Pen never collects:

- Names, email addresses, or account data
- Typed or drawn text content
- Mouse or keyboard recordings
- Hardware identifiers
- Location data

### 5. Data processing location

Telemetry data may be processed on servers outside the European Union. Only anonymous technical data is transmitted.

### 6. Purpose of data processing

Telemetry data is used exclusively for debugging errors, improving application stability, and fixing crashes. No tracking, profiling, or advertising is performed.

### 7. Opt-out

You can disable telemetry at any time in the application settings. Once disabled, no further telemetry data will be sent.

### 8. Contact

If you have questions about privacy or data handling, please open a [GitHub Issue](https://github.com/pythonIsFast/Text2Pen/issues).

## 🤝 Contributing

Contributions are welcome and appreciated.

### How to contribute

1. Fork the repository
2. Create a new branch (`feature/your-feature-name`)
3. Make your changes
4. Commit your changes
5. Open a pull request

### What you can contribute

- Bug fixes
- Performance improvements
- UI/UX enhancements
- Documentation improvements

### Guidelines

- Keep changes focused and minimal
- Write clear commit messages
- Ensure the application still works correctly

For larger changes, please open an issue first to discuss your idea.

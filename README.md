# Text2Pen

Text2Pen converts typed text into handwritten-style mouse input.
It is built for Microsoft OneNote on Windows and also supports Linux-based drawing workflows.

---

## 🖊️ What is Text2Pen?

Text2Pen lets you type text and have it drawn as handwriting strokes.
You train characters once, then reuse your handwriting profile for normal text and tab-separated tables.

---

## ✨ Features

- **Text → handwriting** with your own trained character strokes
- **Table drawing** from tab-separated input
- **Cross-platform binaries** for Windows, Linux x86_64, and Linux arm64
- **AI tools** (chat + quick text actions + optional image input)
- **Automatic updates** with per-asset SHA-256 verification
- **Progress UI** in-app progress + optional always-on-top overlay
- **Optional telemetry** (opt-in crash/error diagnostics)

---

## 🚀 Installation

1. Open the latest [Release](https://github.com/pythonIsFast/Text2Pen/releases/latest).
2. Download the installer asset for your platform:
   - **Windows:** `Installer.exe`
   - **Linux x86_64:** `Installer-linux-x86_64`
   - **Linux arm64:** `Installer-linux-arm64`
3. Run the installer.
4. Start **Text2Pen** from the Start menu / app launcher.

The installer places app files in:
- **Windows:** `%LOCALAPPDATA%\Text2Pen`
- **Linux:** `~/.local/share/Text2Pen`

It also installs the updater as an autostart entry.

### Uninstall

Run the installer again and choose **Uninstall**.

---

## ▶️ Usage

### Windows

1. Open Microsoft OneNote.
2. Start Text2Pen.
3. Enter your text.
4. Click **Start writing** (Text2Pen focuses OneNote and writes into it).

### Linux

1. Open the target app/page where text should be drawn.
2. Start Text2Pen.
3. Enter your text.
4. Click **Start writing**, then pick the start position in the screenshot picker.

---

## ⚠️ Requirements and limitations

- **Windows mode depends on OneNote detection** (a visible OneNote window must be found).
- **Linux writing needs `/dev/uinput` access** (configure udev/group permissions for non-root usage).
- **AI features require internet and backend availability** and may be rate-limited.
- **Input automation can affect other apps** if focus/target position is wrong.

---

## ⚠️ Disclaimer

> **Use Text2Pen at your own risk.**

- Text2Pen automates desktop input and may produce unintended edits.
- Verify target window/position before starting a run.
- Save your work before use.

---

## 🔒 Privacy

- Text2Pen works locally by default.
- Telemetry is **off by default** and only enabled by explicit opt-in.
- Opted-in telemetry is intended for diagnostics (e.g., crash/error reporting).

For questions, open an [issue](https://github.com/pythonIsFast/Text2Pen/issues).

---

## 🤝 Contributing

Contributions are welcome.

1. Fork the repository
2. Create a feature branch
3. Make focused changes
4. Open a pull request

---

## 📜 License

Text2Pen is licensed under the **GNU General Public License v3.0 (GPL-3.0)**.
See [LICENSE](LICENSE).

---

## 🔗 Related projects

- **[Text2Pen Backend](https://github.com/pythonIsFast/Text2Pen-Backend)**

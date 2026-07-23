# Windows Distribution

Build this only on Windows. The application still reads workbook-backed defaults, so the bundled executable must include `MONOFUNI.xls`. Use the checked-in PyInstaller spec instead of a bare `pyinstaller src/winsif_mon/main.py` command.

```powershell
py -m pip install -e .[build]
py -m PyInstaller packaging\winsif_mon.spec
```

The spec includes:

- `MONOFUNI.xls` as bundled data at the executable data root.
- `xlrd` hidden imports, because workbook reading imports it dynamically.
- `packaging/pyinstaller_entry.py` as the entry point, so package-relative imports work in the frozen app.

Do not require the end user to install Python, Excel, VBA tools, or the source tree. Ship the generated `dist\WinSIF_MON` folder as a whole.

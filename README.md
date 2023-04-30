# PTCG Live Auto Redeem Web

This is a Python script for automatically redeeming codes in PTCG Live from website(https://redeem.tcg.pokemon.com/en-us/).

## Dependencies

You will need the following Python packages installed to run the script:

- selenium
- PyInstaller (for building the executable)

You can install these packages using `pip`:

```bash
pip install selenium PyInstaller
```

## Building the Executable

To build the executable, run the following command in your terminal or command prompt:

```bash
pyinstaller --onefile --add-data "Mimikyu.ico;." --icon=big.ico auto_code.py
```

This command will create a standalone executable called auto_code.exe in the dist folder.

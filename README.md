# Excel Value Changer

A simple Windows tool to batch update specific cell values across multiple Excel files.

## Features

- Specify multiple cells at once (e.g., `A1, B2, C3`)
- Batch process all `.xlsx` files in the `data` folder
- Simple GUI interface
- No Excel installation required
- Single executable file

## Screenshot

![Program Screenshot](screenshot.png)

## Usage

### 1. Download
Download `ExcelValueChanger.exe` from [Releases](../../releases).

### 2. Run
1. Place `ExcelValueChanger.exe` in any folder
2. Run the program (a `data` folder will be created automatically)
3. Put your Excel files (.xlsx) in the `data` folder
4. Enter cell addresses and the new value in the program
5. Click "Run Batch Change"

### Cell Address Examples
```
A1
A1, B2, C3
A1, A2, A3, B1, B2, B3
```

## Build from Source

### Requirements
- Python 3.8+

### Build Steps
```bash
# Install dependencies
python -m pip install -r requirements.txt

# Build executable
python -m PyInstaller --onefile --windowed --name "ExcelValueChanger" excel_changer.py
```

Or just run `build_exe.bat` on Windows.

The executable will be created at `dist/ExcelValueChanger.exe`.

## Notes

- Only `.xlsx` files are supported (not `.xls`)
- Back up your files before batch processing
- Close Excel files before running (open files cannot be modified)

## License

MIT License

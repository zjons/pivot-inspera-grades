# pivot-inspera-grades

A small Python script with an optional PyQt5 GUI that pivots and transforms grade export data from Inspera into a friendly Excel sheet. Handles randomized question ordering and produces an easy-to-analyze spreadsheet.

## Features
- Pivot Inspera CSV exports into a per-student grade sheet
- Handles randomized question order gracefully
- CLI script and an optional PyQt5 GUI frontend
- Exports to Excel (.xlsx)
- A Windows executable build is available as a release asset for users who don't want to install Python

## License
This project is released under the GNU Lesser General Public License (LGPL). See the bundled LICENSE file for the full license text.

## Requirements (if running from source)
- Python 3.8+
- pandas
- openpyxl
- PyQt5 (only required for the GUI)

Install dependencies:
```
pip install pandas openpyxl PyQt5
```

## Usage

### CLI
Run the pivot script from the command line:
```
python pivot_inspera_grades.py input.csv output.xlsx
```
Replace the filenames with your actual Inspera export and the desired output file.

### GUI
To run the GUI:
```
python pivot_inspera_grades_gui.py
```
The GUI will prompt you to select the input file and output location.

### Windows executable (recommended if you don't want to install Python)
A pre-built Windows executable is provided as a release asset. Download the latest release from:
https://github.com/zjons/pivot-inspera-grades/releases

After downloading the .exe, run it and use the GUI to select your Inspera CSV and export location.

## Building the Windows executable
The repository contains a build script to create the Windows executable. If you prefer to build the .exe yourself (for example to inspect or tweak the bundled code), run the provided build script from a Windows environment.

## Example
- Input: Inspera CSV export (rows are attempts, columns include randomized question answers)
- Output: Excel workbook with pivoted grades per student and any summary sheets the script creates

## Contributing
Contributions are welcome. Please open an issue or submit a PR with a clear description of the change. If you open an issue about a bug or feature request, including a sanitized sample Inspera export will help reproduction.

## Notes
- If you have a sample Inspera export (sanitized), include it when creating issues so I can reproduce and test changes.

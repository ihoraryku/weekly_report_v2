# Solar Forecast Report Generator

This application generates weekly solar energy forecast reports using template files and source data. It is built using PyQt5 for the user interface and OpenPyXL for handling Excel files.

## Features

- Select directories for templates, source data, and output.
- Generate reports for a specified week.
- Parallel processing of templates to improve performance.
- Displays progress and elapsed time during report generation.
- Alerts if any source files are missing for the selected week.

## Requirements

- Python 3.x
- PyQt5
- OpenPyXL

## Installation

1. Clone the repository:
   ```bash
   git clone <repository-url>
   cd <repository-directory>
   ```

2. Install the required packages:
   ```bash
   pip install -r requirements.txt
   ```

## Usage

1. Run the application:
   ```bash
   python report.py
   ```

2. Use the interface to select the directories for templates, source data, and output.

3. Select the start date for the report using the calendar widget.

4. Click "Згенерувати звіт" to generate the report.

5. The application will display progress and notify you when the report generation is complete.

## Creating an Executable with PyInstaller

To create a standalone executable file for the application, you can use PyInstaller. This allows you to distribute the application without requiring users to have Python installed.

1. Install PyInstaller:
   ```bash
   pip install pyinstaller
   ```

2. Generate the executable:
   ```bash
   pyinstaller --onefile --windowed report.py
   ```

   - `--onefile`: Creates a single executable file.
   - `--windowed`: Suppresses the console window (useful for GUI applications).

3. The executable will be located in the `dist` directory.

## Logging

- Logs are saved to `app.log` in the project directory.
- Logs include information about the report generation process and any errors encountered.

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.

## Contributing

Contributions are welcome! Please fork the repository and submit a pull request with your changes.

## Contact

For questions or feedback, please contact Ihor Aryku at igorarycu@gmail.com.
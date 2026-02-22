# CyberCity Automation

An automated data extraction tool for the CyberCity application designed to facilitate scheduled vehicle data retrieval and summary generation. This project uses GUI automation to interface with the CyberCity desktop application, allowing for systematic extraction of vehicle records over specified time intervals.

## Features

- Automated GUI interaction for camera selection and data filtering.
- Interval-based data extraction (e.g., 10-minute segments).
- Batch processing for historical data retrieval (e.g., weekly extracts).
- Automated Excel summary generation from multiple extracted files.
- Interactive UI calibration system for adaptable screen coordinates.
- Error handling with automated screenshot capture on failure.

## Prerequisites

- Python 3.x
- Tesseract OCR (if using OCR features)
- CyberCity desktop application

## Installation

1. Clone the repository:

   ```bash
   git clone <repository-url>
   cd cyber
   ```

2. Install the required Python dependencies:

   ```bash
   pip install -r requirements.txt
   ```

## Configuration and Calibration

Before running the automation, you must calibrate the UI coordinates to match your display settings and application layout.

1. Open the CyberCity application and ensure it is visible on your screen.
2. Run the calibration process:

   ```python
   python automate.py
   ```

   *Note: The script will prompt you to position your mouse over specific UI elements (dropdowns, input fields, buttons). Follow the on-screen instructions.*

3. The coordinates will be saved to `cybercity_coordinates.json` for future use.

## Usage

### Basic Execution

Modify the `CAMERA_NAME` and `OUTPUT_FOLDER` variables in `automate.py` (or via a separate configuration script), then execute:

```bash
python automate.py
```

### Scripted Automation

You can also use the `CyberCityAutomation` class in your own scripts:

```python
from automate import CyberCityAutomation
from datetime import datetime, timedelta

# Initialize the automator
automator = CyberCityAutomation(camera_name="Camera_01", output_folder="data_exports")

# Define time range
start = datetime(2023, 10, 1, 8, 0)
end = datetime(2023, 10, 1, 18, 0)

# Extract data in 15-minute intervals
automator.extract_week_in_intervals(start, end, interval_minutes=15)
```

## Data Output

Extracted data is saved as `.xlsx` files in the specified output folder. Upon completion of a batch extraction, a summary file (e.g., `summary_YYYYMMDD_to_YYYYMMDD.xlsx`) is automatically generated, consolidating all records into a single spreadsheet.

## Project Structure

- `automate.py`: Core automation logic and GUI interaction class.
- `config.py`: Example script for initializing the automation.
- `requirements.txt`: Project dependencies.
- `cybercity_coordinates.json`: (Generated) UI element coordinates.

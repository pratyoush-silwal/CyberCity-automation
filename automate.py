import pyautogui
import time
import pandas as pd
from datetime import datetime, timedelta
import os
import subprocess
import sys
import pytesseract
from PIL import Image
import cv2
import numpy as np


class CyberCityAutomation:
    def __init__(self, camera_name, output_folder):
        """
        Initialize the automation class

        Args:
            camera_name: Name of the camera as it appears in dropdown
            output_folder: Folder where Excel files will be saved
        """
        self.camera_name = camera_name
        self.output_folder = output_folder

        # Create output folder if it doesn't exist
        if not os.path.exists(output_folder):
            os.makedirs(output_folder)

        # Confidence level for image recognition (adjust if needed)
        pyautogui.PAUSE = 1  # Add 1 second pause between actions
        pyautogui.FAILSAFE = True  # Move mouse to top-left corner to abort

        # Store coordinates for important UI elements (you'll need to calibrate these)
        self.ui_coordinates = self.calibrate_ui()

    def calibrate_ui(self):
        """
        Interactive calibration to get UI element positions
        Run this once to set up coordinates for your specific screen setup
        """
        print("=== UI Calibration Mode ===")
        print("You have 5 seconds to position your mouse over each element")
        print("Press Ctrl+C to skip calibration")

        coordinates = {}
        elements = [
            "camera_dropdown",
            "license_plate_field",
            "start_date_field",
            "start_time_field",
            "end_date_field",
            "end_time_field",
            "extract_button",
            "save_dialog",
        ]

        try:
            for element in elements:
                input(f"\nPosition mouse over {element} and press Enter...")
                time.sleep(2)
                coordinates[element] = pyautogui.position()
                print(f"✓ {element} captured at {coordinates[element]}")

            # Save coordinates to file
            import json

            with open("cybercity_coordinates.json", "w") as f:
                json.dump({k: [v.x, v.y] for k, v in coordinates.items()}, f)
            print("\n✓ Coordinates saved to cybercity_coordinates.json")

        except KeyboardInterrupt:
            print("\nCalibration skipped, using default coordinates")
            # You'll need to update these with actual values after calibration
            coordinates = {
                "camera_dropdown": (100, 100),
                "license_plate_field": (200, 150),
                "start_date_field": (300, 200),
                "start_time_field": (400, 200),
                "end_date_field": (500, 200),
                "end_time_field": (600, 200),
                "extract_button": (700, 300),
                "save_dialog": (400, 400),
            }

        return coordinates

    def launch_app(self, app_path=None):
        """Launch the CyberCity application if not already running"""
        if app_path:
            subprocess.Popen([app_path])
            time.sleep(10)  # Wait for app to load

    def select_camera(self):
        """Select camera from dropdown menu"""
        # Click on camera dropdown
        pyautogui.click(self.ui_coordinates["camera_dropdown"])
        time.sleep(1)

        # Type camera name (assuming dropdown has search)
        pyautogui.write(self.camera_name)
        time.sleep(1)

        # Press Enter to select
        pyautogui.press("enter")
        time.sleep(2)

    def clear_license_plate(self):
        """Clear license plate search field if it contains text"""
        # Click on license plate field
        pyautogui.click(self.ui_coordinates["license_plate_field"])
        time.sleep(0.5)

        # Select all text and delete
        pyautogui.hotkey("ctrl", "a")
        pyautogui.press("delete")
        time.sleep(0.5)

    def set_datetime(self, field_type, datetime_value):
        """
        Set date or time field

        Args:
            field_type: 'start_date', 'start_time', 'end_date', or 'end_time'
            datetime_value: datetime object or string
        """
        coordinate_key = f"{field_type}_field"

        # Click on the field
        pyautogui.click(self.ui_coordinates[coordinate_key])
        time.sleep(0.5)

        # Select all existing text
        pyautogui.hotkey("ctrl", "a")
        pyautogui.press("delete")

        # Format the value based on field type
        if isinstance(datetime_value, datetime):
            if "date" in field_type:
                value = datetime_value.strftime("%Y-%m-%d")  # Adjust format as needed
            else:
                value = datetime_value.strftime("%H:%M")
        else:
            value = str(datetime_value)

        # Type the value
        pyautogui.write(value)
        pyautogui.press("enter")
        time.sleep(1)

    def click_extract(self):
        """Click the extract button"""
        pyautogui.click(self.ui_coordinates["extract_button"])
        time.sleep(5)  # Wait for extraction to complete

    def save_file(self, filename):
        """
        Save the extracted file

        Args:
            filename: Name for the saved file
        """
        # Wait for save dialog
        time.sleep(2)

        # Click on save dialog (adjust based on your app)
        pyautogui.click(self.ui_coordinates["save_dialog"])
        time.sleep(1)

        # Type filename
        pyautogui.write(filename)
        time.sleep(1)

        # Press Enter to save
        pyautogui.press("enter")
        time.sleep(3)  # Wait for save to complete

    def extract_interval(self, start_time, end_time):
        """
        Extract data for a specific time interval

        Args:
            start_time: datetime object for start
            end_time: datetime object for end
        """
        # Clear license plate field
        self.clear_license_plate()

        # Set start date and time
        self.set_datetime("start_date", start_time)
        self.set_datetime("start_time", start_time)

        # Set end date and time
        self.set_datetime("end_date", end_time)
        self.set_datetime("end_time", end_time)

        # Generate filename
        filename = f"vehicles_{start_time.strftime('%Y%m%d_%H%M')}_to_{end_time.strftime('%Y%m%d_%H%M')}.xlsx"
        filepath = os.path.join(self.output_folder, filename)

        # Click extract
        self.click_extract()

        # Save file
        self.save_file(filepath)

        print(f"✓ Extracted: {filename}")

        return filepath

    def extract_week_in_intervals(self, start_date, end_date, interval_minutes=10):
        """
        Extract a full week in specified intervals

        Args:
            start_date: Start date (datetime)
            end_date: End date (datetime)
            interval_minutes: Interval length in minutes
        """
        current_time = start_date
        total_intervals = 0
        successful_intervals = 0

        # Calculate total intervals
        total_minutes = int((end_date - start_date).total_seconds() / 60)
        total_intervals = total_minutes // interval_minutes

        print(f"Starting extraction from {start_date} to {end_date}")
        print(f"Total intervals: {total_intervals}")

        while current_time < end_date:
            interval_end = min(
                current_time + timedelta(minutes=interval_minutes), end_date
            )

            try:
                # Select camera (in case selection was lost)
                self.select_camera()

                # Extract this interval
                filepath = self.extract_interval(current_time, interval_end)

                # Verify file was created
                if os.path.exists(filepath):
                    successful_intervals += 1
                    print(f"Progress: {successful_intervals}/{total_intervals}")
                else:
                    print(f"⚠ File not created for interval {current_time}")

                # Small delay between extractions
                time.sleep(2)

            except Exception as e:
                print(f"✗ Error at interval {current_time}: {str(e)}")

                # Optional: Take screenshot on error
                screenshot = pyautogui.screenshot()
                screenshot.save(f"error_{current_time.strftime('%Y%m%d_%H%M')}.png")

            # Move to next interval
            current_time = interval_end

        print(f"\n✓ Extraction complete!")
        print(
            f"Successfully extracted {successful_intervals} out of {total_intervals} intervals"
        )

        # Create summary Excel file
        self.create_summary_file(start_date, end_date, interval_minutes)

    def create_summary_file(self, start_date, end_date, interval_minutes):
        """Create a summary Excel file with all extracted data"""
        print("\nCreating summary file...")

        all_data = []

        # Get all extracted files
        files = sorted(
            [
                f
                for f in os.listdir(self.output_folder)
                if f.startswith("vehicles_") and f.endswith(".xlsx")
            ]
        )

        for file in files:
            filepath = os.path.join(self.output_folder, file)
            try:
                df = pd.read_excel(filepath)
                # Add interval information
                df["Source_File"] = file
                all_data.append(df)
            except Exception as e:
                print(f"Could not read {file}: {str(e)}")

        if all_data:
            combined_df = pd.concat(all_data, ignore_index=True)
            summary_file = os.path.join(
                self.output_folder,
                f"summary_{start_date.strftime('%Y%m%d')}_to_{end_date.strftime('%Y%m%d')}.xlsx",
            )
            combined_df.to_excel(summary_file, index=False)
            print(f"✓ Summary file created: {summary_file}")
            print(f"Total records: {len(combined_df)}")
        else:
            print("No data files found to summarize")


def main():
    """Main execution function"""

    # Configuration
    CAMERA_NAME = "Your Camera Name Here"  # Replace with actual camera name
    OUTPUT_FOLDER = "extracted_vehicle_data"

    # Set date range (example: last 7 days)
    end_date = datetime.now().replace(hour=23, minute=59, second=59)
    start_date = (end_date - timedelta(days=7)).replace(hour=0, minute=0, second=1)

    # Initialize automation
    automator = CyberCityAutomation(CAMERA_NAME, OUTPUT_FOLDER)

    # Optional: Launch app (uncomment and set path if needed)
    # automator.launch_app("C:\\Path\\To\\CyberCity.exe")

    print("=" * 50)
    print("CyberCity Automation Started")
    print(f"Camera: {CAMERA_NAME}")
    print(f"Date Range: {start_date} to {end_date}")
    print("=" * 50)

    # Give user time to focus the app window
    print("\nSwitch to CyberCity app window now!")
    for i in range(5, 0, -1):
        print(f"Starting in {i} seconds...")
        time.sleep(1)

    try:
        # Select camera initially
        automator.select_camera()

        # Extract data in 10-minute intervals
        automator.extract_week_in_intervals(start_date, end_date, interval_minutes=10)

    except KeyboardInterrupt:
        print("\n\nAutomation stopped by user")
    except Exception as e:
        print(f"\nUnexpected error: {str(e)}")
        import traceback

        traceback.print_exc()

    print("\n✓ Automation finished")


if __name__ == "__main__":
    # Install required packages if not present
    required_packages = ["pyautogui", "pandas", "openpyxl", "Pillow"]

    for package in required_packages:
        try:
            __import__(package)
        except ImportError:
            print(f"Installing {package}...")
            subprocess.check_call([sys.executable, "-m", "pip", "install", package])

    main()

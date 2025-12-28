"""
Stock Market Data Fetch - Executable Build Script

This script creates an executable (.exe) file from the Stock Market Data application.
"""

import os
import sys
import subprocess
import shutil
import time
from datetime import datetime

def remove_directory_with_retry(path, max_retries=5, delay=1):
    """Remove directory with retry logic for Windows file locking issues."""
    for attempt in range(max_retries):
        try:
            if os.path.exists(path):
                shutil.rmtree(path)
            return True
        except (PermissionError, OSError) as e:
            if attempt < max_retries - 1:
                print(f"      Retry {attempt + 1}/{max_retries - 1} (waiting {delay}s...)")
                time.sleep(delay)
            else:
                return False
    return True

def create_executable():
    """Create an executable file from the application."""
    print("\n" + "="*60)
    print("STOCKWATCH - EXECUTABLE BUILD SCRIPT")
    print("="*60)

    # Record the start time
    start_time = datetime.now()
    print(f"Build started at: {start_time.strftime('%H:%M:%S')}")
    print("Expected duration: 3-5 minutes")

    # Clean up old build artifacts
    print("\nCleaning up old build artifacts...")
    if os.path.exists("build"):
        print("  - Removing old build/ directory...")
        if remove_directory_with_retry("build"):
            print("      ✓ Removed build/")
        else:
            print("      ⚠ Warning: Could not remove build/ (files may be in use)")
            print("      The build will continue anyway...")

    if os.path.exists("dist"):
        print("  - Removing old dist/ directory...")
        if remove_directory_with_retry("dist"):
            print("      ✓ Removed dist/")
        else:
            print("      ✗ ERROR: Could not remove dist/ after multiple retries")
            print("      ")
            print("      Possible solutions:")
            print("      1. Check Task Manager for StockMarketApp.exe and end it")
            print("      2. Close all File Explorer windows")
            print("      3. Disable antivirus temporarily")
            print("      4. Manually delete the dist folder and try again")
            print("      5. Restart your computer if the issue persists")
            return False

    # Create necessary directories
    print("\nPreparing build environment...")
    print("  - Creating build directories...")
    os.makedirs("build", exist_ok=True)
    print("      ✓ build/")
    os.makedirs("dist", exist_ok=True)
    print("      ✓ dist/")
    os.makedirs("yahoo_finance_data", exist_ok=True)
    print("      ✓ yahoo_finance_data/")

    # Check if EQUITY_L.csv exists before building
    csv_src = "EQUITY_L.csv"
    print(f"\n  - Checking for required files...")
    if not os.path.exists(csv_src):
        print(f"      ✗ ERROR: {csv_src} file not found in the project directory!")
        print("Make sure the file exists before building the executable.")
        return False
    print(f"      ✓ {csv_src} found")
    
    # Use python -m PyInstaller instead of direct pyinstaller command
    pyinstaller_cmd = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--name=StockMarketApp",
        "--onedir",                    # Create a directory containing the executable
        "--windowed",                  # GUI application (no console)
        f"--add-data={csv_src}{os.pathsep}.",   # Include the stock symbols CSV file with OS-specific separator
        "--noconfirm",                 # Overwrite output directory if it exists
        "--clean",                     # Clean PyInstaller cache and temporary files
        "--log-level=WARN",            # Reduce log verbosity
        # Exclude unnecessary heavy packages to speed up build
        "--exclude-module=torch",
        "--exclude-module=tensorflow",
        "--exclude-module=tensorboard",
        "--exclude-module=scipy",
        "--exclude-module=matplotlib.tests",
        "--exclude-module=numpy.tests",
        "--exclude-module=pandas.tests",
        "--exclude-module=IPython",
        "--exclude-module=jupyter",
        "--exclude-module=notebook",
        "--exclude-module=sphinx",
        "--exclude-module=pytest",
        "--exclude-module=PIL.tests",
        "main.py"                      # Entry point script
    ]
    
    # Execute PyInstaller
    print("\n" + "="*60)
    print("STEP 1/3: Running PyInstaller to build the executable...")
    print("="*60)
    print("This is the longest step and may take 3-5 minutes...")
    print(f"Command: {' '.join(pyinstaller_cmd)}")
    print("\nPyInstaller output:\n")

    pyinstaller_start = datetime.now()

    try:
        # Run without capturing output so we can see progress in real-time
        result = subprocess.run(pyinstaller_cmd)

        pyinstaller_end = datetime.now()
        pyinstaller_duration = pyinstaller_end - pyinstaller_start

        # Check if PyInstaller was successful
        if result.returncode != 0:
            print("\n" + "="*60)
            print("ERROR: PyInstaller failed!")
            print("="*60)
            return False
        else:
            print("\n" + "="*60)
            print(f"PyInstaller completed successfully in {pyinstaller_duration}")
            print("="*60)
    except Exception as e:
        print(f"\nError running PyInstaller: {str(e)}")
        return False
    
    # Copy additional necessary files
    print("\n" + "="*60)
    print("STEP 2/3: Copying additional files to distribution folder...")
    print("="*60)

    copy_start = datetime.now()

    try:
        # Manually copy EQUITY_L.csv to ensure it's in the right place
        dist_folder = os.path.join("dist", "StockMarketApp")
        print(f"[1/3] Verifying distribution folder: {dist_folder}")

        # Copy CSV file (double check)
        csv_dst = os.path.join(dist_folder, "EQUITY_L.csv")
        print(f"[2/3] Copying {csv_src} to distribution folder...")
        if os.path.exists(csv_src):
            shutil.copy2(csv_src, csv_dst)
            print(f"      ✓ Successfully copied {csv_src}")
        else:
            print(f"      WARNING: Could not find {csv_src} to copy!")

        # Create watchlist.txt if it doesn't exist
        watchlist_path = os.path.join(dist_folder, "watchlist.txt")
        print(f"[3/3] Creating empty watchlist.txt...")
        with open(watchlist_path, "w") as f:
            pass
        print(f"      ✓ Created watchlist.txt")

        # Create the yahoo_finance_data directory in the distribution folder
        yahoo_dir = os.path.join(dist_folder, "yahoo_finance_data")
        os.makedirs(yahoo_dir, exist_ok=True)
        print(f"      ✓ Created yahoo_finance_data directory")
        
        copy_end = datetime.now()
        copy_duration = copy_end - copy_start
        print(f"\n      File operations completed in {copy_duration}")

        # Verify that critical files exist in the distribution
        print("\n" + "="*60)
        print("STEP 3/3: Verifying build output...")
        print("="*60)

        verify_start = datetime.now()

        if not os.path.exists(csv_dst):
            print(f"      WARNING: Could not find {csv_src} in the distribution folder!")
            print(f"      Trying again with absolute paths...")
            # Try with absolute paths as a fallback
            abs_src = os.path.abspath(csv_src)
            abs_dst = os.path.abspath(csv_dst)
            print(f"      Copying {abs_src} to {abs_dst}...")
            shutil.copy2(abs_src, abs_dst)

            if os.path.exists(csv_dst):
                print(f"      ✓ Successfully copied {csv_src} on second attempt")
            else:
                print(f"      ERROR: Failed to copy {csv_src} to the distribution folder!")
                return False

        # Verify all critical files
        exe_path = os.path.join(dist_folder, "StockMarketApp.exe")
        print(f"\nVerifying critical files:")
        print(f"      {'✓' if os.path.exists(exe_path) else '✗'} StockMarketApp.exe")
        print(f"      {'✓' if os.path.exists(csv_dst) else '✗'} EQUITY_L.csv")
        print(f"      {'✓' if os.path.exists(watchlist_path) else '✗'} watchlist.txt")
        print(f"      {'✓' if os.path.exists(yahoo_dir) else '✗'} yahoo_finance_data/")

        verify_end = datetime.now()
        verify_duration = verify_end - verify_start
        print(f"\n      Verification completed in {verify_duration}")

    except Exception as e:
        print(f"\nError copying additional files: {str(e)}")
        return False

    # Record end time and calculate duration
    end_time = datetime.now()
    duration = end_time - start_time

    print("\n" + "="*60)
    print("BUILD SUCCESSFUL!")
    print("="*60)
    print(f"Build started:    {start_time.strftime('%H:%M:%S')}")
    print(f"Build completed:  {end_time.strftime('%H:%M:%S')}")
    print(f"Total build time: {duration}")
    print("\nOutput location:")
    print(f"  Folder: dist/StockMarketApp/")
    print(f"  Executable: dist/StockMarketApp/StockMarketApp.exe")
    print("="*60)
    
    return True

if __name__ == "__main__":
    create_executable()
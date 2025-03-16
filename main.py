import subprocess
import os
import time

# Define the paths
pdf_path = r"D:\Kali\Usman Mustafa Khawer.pdf"
acrobat_path = r"C:\Program Files (x86)\Adobe\Reader 11.0\Reader\AcroRd32.exe"

# Check if the Adobe Reader path exists
if os.path.exists(acrobat_path):
    # Command for Adobe Reader to print
    command = [acrobat_path, '/t', pdf_path]
else:
    # Fallback to SumatraPDF if Adobe Reader is not found
    sumatra_path = "SumatraPDF-3.5.2-64.exe"
    command = [sumatra_path, '-print-to-default', pdf_path]

# Start the subprocess
try:
    process = subprocess.Popen(command)

    # Check every second if the process is still running
    while True:
        # Check if the process is still alive
        retcode = process.poll()  # Returns None if process is still running
        if retcode is not None:
            # Process finished
            print("Adobe Reader process completed.")
            break

        # Wait for 1 second before checking again
        time.sleep(1)

    # After process has finished, kill any remaining Adobe Reader processes
    subprocess.call(["taskkill", "/F", "/IM", "AcroRd32.exe"])
    print("Adobe Reader process killed successfully.")

except Exception as e:
    print(f"An error occurred: {e}")

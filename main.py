import os
import requests
import PyPDF2
import threading
import queue
from datetime import datetime

import win32api
import win32print


VOUCHER_CONTAINER = {}
ALREADY_PRINTED_RFID = []
TEMP_URL = ""
GHOSTSCRIPT_PATH = "C:/Program Files (x86)/gs/gs9.55.0/bin/gswin32.exe"
GSPRINT_PATH = ""

PROXY_URL = '172.16.17.9'
PROXY_PORT = '8080'

PROXIES = {
    'http': f'http://{PROXY_URL}:{PROXY_PORT}',
    'https': f'http://{PROXY_URL}:{PROXY_PORT}',
}


input_queue = queue.Queue()


def download(url: str, dest_folder: str, id: str):
    if not os.path.exists(dest_folder):
        os.makedirs(dest_folder)  # create folder if it does not exist

    filename = id  # be careful with file names
    file_path = os.path.join(dest_folder, filename)

    r = requests.get(url, stream=True, proxies=PROXIES)
    if r.ok:
        # print("saving to", os.path.abspath(file_path))
        print("Fetching Voucher... " + filename)
        with open(file_path, 'wb') as f:
            for chunk in r.iter_content(chunk_size=1024 * 8):
                if chunk:
                    f.write(chunk)
                    f.flush()
                    os.fsync(f.fileno())
    else:  # HTTP status code 4XX/5XX
        print("Download failed: status code {}\n{}".format(r.status_code, r.text))


def fetching_data():
    current_date = datetime.now()
    today_date = current_date.strftime("%d-%b-%y")

    global GHOSTSCRIPT_PATH, GSPRINT_PATH
    with open("ghostscript_path.ini", "r") as file:
        for line in file:
            value = line.strip()
            if value.__contains__("GSPRINT_PATH = "):
                GSPRINT_PATH = value.split("GSPRINT_PATH = ")[1]

    existing_txt_files = [file for file in os.listdir() if file.endswith(".txt")]

    if not os.path.exists(today_date + ".txt"):
        with open(today_date + ".txt", "w") as file:
            # Write the formatted date to the file
            pass

    for file in existing_txt_files:
        if file != f"{today_date}.txt":
            os.remove(file)
            print(f"Deleted: {file}")
        else:
            with open(file, "r") as file:
                for line in file:
                    ALREADY_PRINTED_RFID.append(line.strip())

    voucher_google_sheet = 'https://script.google.com/macros/s/AKfycbztIlguZFMWj3fNwIwUpIG7XDyBxB2u2b4Li6HZWwCIsPHbs3eMxJrKM_BfaAXexQC8/exec'
    today_report_url = ""
    voucher_container = {}

    response = requests.get(voucher_google_sheet, proxies=PROXIES)
    if response.status_code == 200:
        json_data = response.json()
        today_report_url = json_data['URL']
        for i in json_data['data']:
            if i['student_id_barcode'] not in voucher_container:
                voucher_container[str(i['student_id_barcode'])] = [i['student_id'], str(i['student_voucher'])]

    global TEMP_URL, VOUCHER_CONTAINER
    TEMP_URL = today_report_url
    VOUCHER_CONTAINER = voucher_container


def find_and_print(rfid):
    current_date = datetime.now()
    today_date = current_date.strftime("%d-%b-%y")

    global GHOSTSCRIPT_PATH, GSPRINT_PATH

    if rfid in VOUCHER_CONTAINER:
        if str(rfid) not in ALREADY_PRINTED_RFID:
            ALREADY_PRINTED_RFID.append(str(rfid))
            with open(today_date + ".txt", "a") as file:
                # Append a new line with a value
                new_value = str(rfid)
                file.write(new_value + "\n")

            url = TEMP_URL.split('student_id=')[0] + "student_id=" + VOUCHER_CONTAINER[rfid][0] + "&v_voucher_no=" + VOUCHER_CONTAINER[rfid][1]

            file_name = 'temp.pdf'
            file_location = "b/"
            download(url, file_location, file_name)

            input_path = file_location + file_name  # Replace with your input PDF file path
            rotation_angle = 90  # Specify the rotation angle (90 degrees in this example)

            with open(input_path, 'rb+') as file:
                pdf_reader = PyPDF2.PdfFileReader(file)
                pdf_writer = PyPDF2.PdfFileWriter()

                # Process only the first page
                if pdf_reader.numPages > 0:
                    page = pdf_reader.getPage(0)
                    page.rotateClockwise(rotation_angle)
                    pdf_writer.addPage(page)

                file.seek(0)  # Move the file cursor to the beginning before writing
                pdf_writer.write(file)
                file.truncate()  # Remove any remaining

            # YOU CAN PUT HERE THE NAME OF YOUR SPECIFIC PRINTER INSTEAD OF DEFAULT
            currentprinter = win32print.GetDefaultPrinter()


            # Construct the Ghostscript command with the -dOrientation=3 option for landscape mode
            ghostscript_cmd = f'-ghostscript "{GHOSTSCRIPT_PATH}" -printer "{currentprinter}" -dOrientation=1 "{os.path.abspath(file_location + file_name)}"'

            win32api.ShellExecute(0, 'open', GSPRINT_PATH, ghostscript_cmd, '.', 0)
            print("Voucher Print Successfully\n")
        else:
            print("Voucher Already Printed\n")
    else:
        print("No Voucher Found\n")


def process_worker():
    while True:
        if not input_queue.empty():
            value = input_queue.get()
            print("")
            find_and_print(value)


def main():
    print("Fetching Voucher Data....\n\n")
    fetching_data()
    # Start the process function in a separate thread
    process_thread = threading.Thread(target=process_worker)
    process_thread.start()

    while True:
        value = input("Punch Card Here:")
        # Enqueue the input
        input_queue.put(value)


if __name__ == "__main__":
    main()
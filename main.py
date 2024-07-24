import os
import requests
import urllib.request

import PyPDF2
import threading
import queue
from datetime import datetime

import win32api
import win32print



VOUCHER_CONTAINER = {}
ALREADY_PRINTED_RFID = []
TEMP_URL = ""
GHOSTSCRIPT_PATH = 'C:/Program Files (x86)/gs/gs9.55.0/bin/gswin32c.exe'
GSPRINT_PATH = ""

# PROXY_URL = '172.16.17.9'
# PROXY_PORT = '8080'
#
# PROXIES = {
#     'http': f'http://{PROXY_URL}:{PROXY_PORT}',
#     'https': f'http://{PROXY_URL}:{PROXY_PORT}',
# }


input_queue = queue.Queue()


def print_job_checker():
    default_printer = win32print.GetDefaultPrinter()
    jobs = [1]
    document = "No Document"

    while jobs:
        jobs = []
        phandle = win32print.OpenPrinter(default_printer)
        try:
            print_jobs = win32print.EnumJobs(phandle, 0, -1, 1)
            if print_jobs:
                jobs.extend(list(print_jobs))
            for job in print_jobs:
                document = job["pDocument"]
        finally:
            win32print.ClosePrinter(phandle)

        if not jobs:
            return None
        else:
            return str(document).split("\\b\\")[1].split(".pdf")[0] + " Printing Voucher..."


def download(url: str, dest_folder: str, id: str):
    if not os.path.exists(dest_folder):
        os.makedirs(dest_folder)  # create folder if it does not exist

    filename = id  # be careful with file names
    file_path = os.path.join(dest_folder, filename)

    # Request the content without streaming
    try:
        # Download the file from the URL
        with urllib.request.urlopen(url) as response, open(file_path, 'wb') as out_file:
            data = response.read()  # Read the response data
            out_file.write(data)  # Write the data to a file
    except urllib.error.URLError as e:
        print(f"Failed to download file: {e.reason}")


def fetching_data():
    current_date = datetime.now()
    today_date = current_date.strftime("%d-%b-%y")

    with open("ghostscript_path.ini", "r") as file:
        for line in file:
            value = line.strip()
            if value.__contains__("GSPRINT_PATH = "):
                GSPRINT_PATH = value.split("GSPRINT_PATH = ")[1]

    existing_txt_files = [file for file in os.listdir() if file.endswith(".txt")]

    if not os.path.exists(today_date + ".txt"):
        with open(today_date + ".txt", "w") as file:
            pass

    for file in existing_txt_files:
        if file != f"{today_date}.txt":
            os.remove(file)
            print(f"Deleted: {file}")
        else:
            with open(file, "r") as file:
                for line in file:
                    ALREADY_PRINTED_RFID.append(line.strip())

    voucher_google_sheet = 'https://script.googleusercontent.com/macros/echo?user_content_key=70D3O3UTxkZiyeUjCTUSiH4fEYQFpyHPwUba8GMgXsYes2LSR07KcQUHarTuA0TKsY4SvlDuLz4Vjg-JrBb2lYA2eitMfN0lm5_BxDlH2jW0nuo2oDemN9CCS2h10ox_1xSncGQajx_ryfhECjZEnM1YqgwiDg8ywLFXEeGOTMOs8f-3eoR-8izo93E2rchqKAlp4XmWTBlIed5Uhlj9wQQJ7Sq0DTH49ELx3XxYzUQ6gFOLIEm9ow&lib=M8MTx88xD5ZuMYQh6PpftO3O2iu_L2FkJ'
    today_report_url = ""
    voucher_container = {}

    # For Proxy
    # response = requests.get(voucher_google_sheet, proxies=PROXIES)
    response = requests.get(voucher_google_sheet)
    if response.status_code == 200:
        json_data = response.json()
        today_report_url = json_data['URL']
        for i in json_data['data']:
            if i['student_id_barcode'] not in voucher_container:
                voucher_container[str(i['student_id_barcode'])] = [i['student_id'], i['student_name'], str(i['student_voucher'])]

    global TEMP_URL, VOUCHER_CONTAINER
    TEMP_URL = today_report_url
    VOUCHER_CONTAINER = voucher_container


def find_and_print(rfid):
    if print_job_checker():
        print(print_job_checker())
    current_date = datetime.now()
    today_date = current_date.strftime("%d-%b-%y")

    global GHOSTSCRIPT_PATH, GSPRINT_PATH
    with open("ghostscript_path.ini", "r") as file:
        for line in file:
            value = line.strip()
            if value.__contains__("GSPRINT_PATH = "):
                GSPRINT_PATH = value.split("GSPRINT_PATH = ")[1]

    if rfid in VOUCHER_CONTAINER:
        if str(rfid) not in ALREADY_PRINTED_RFID:
            ALREADY_PRINTED_RFID.append(str(rfid))
            with open(today_date + ".txt", "a") as file:
                # Append a new line with a value
                new_value = str(rfid)
                file.write(new_value + "\n")

            url = TEMP_URL.split('student_id=')[0] + "student_id=" + VOUCHER_CONTAINER[rfid][0] + "&v_voucher_no=" + VOUCHER_CONTAINER[rfid][2]

            file_name = VOUCHER_CONTAINER[rfid][1] + '.pdf'
            file_location = "b/"

            try:
                download(url, file_location, file_name)

                # input_path = file_location + file_name
                # rotation_angle = 90
                #
                # with open(input_path, 'rb+') as file:
                #     pdf_reader = PyPDF2.PdfFileReader(file)
                #     pdf_writer = PyPDF2.PdfFileWriter()
                #
                #     # Process only the first page
                #     if pdf_reader.numPages > 0:
                #         page = pdf_reader.getPage(0)
                #         page.rotateClockwise(rotation_angle)
                #         pdf_writer.addPage(page)
                #
                #     file.seek(0)
                #     pdf_writer.write(file)
                #     file.truncate()

                currentprinter = win32print.GetDefaultPrinter()
                ghostscript_cmd = f'-ghostscript "{GHOSTSCRIPT_PATH}" -printer "{currentprinter}" -dOrientation=1 "{os.path.abspath(file_location + file_name)}"'

                win32api.ShellExecute(0, 'open', GSPRINT_PATH, ghostscript_cmd, '.', 0)

                print("Voucher Print Successfully\n")
            except Exception as e:
                with open(today_date + ".txt", 'r') as file:
                    lines = file.readlines()

                lines = lines[:-1]
                with open(today_date + ".txt", 'w') as file:
                    file.writelines(lines)

                if rfid in ALREADY_PRINTED_RFID:
                    ALREADY_PRINTED_RFID.remove(rfid)

                print("Failed To Download Voucher... " + e)

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
        if print_job_checker():
            print(print_job_checker())


def main():
    print("Fetching Voucher Data....\n\n")
    print("Printer State Ready....\n\n")
    fetching_data()
    # Start the process function in a separate thread
    process_thread = threading.Thread(target=process_worker)
    process_thread.start()

    while True:
        value = input("")
        # Enqueue the input
        # input_queue.put(value)
        find_and_print(value)


if __name__ == "__main__":
    main()
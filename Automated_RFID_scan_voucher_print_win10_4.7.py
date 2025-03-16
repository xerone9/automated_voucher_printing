import customtkinter
from tkinter import *
from tkinter import ttk
from PIL import Image
import time

import os
import requests
import urllib.request
import urllib.error

import threading
import queue
from datetime import datetime

import win32api
import win32print
import win32timezone

from get_voucher_from_database import get_pending_vouchers

proxies = {
    "http": "http://172.16.0.9:8080",
    "https": "http://172.16.0.9:8080"
}

input_queue = queue.Queue()
kuickpay_queue = queue.Queue()


ALREADY_PRINTED_VOUCHER = []
KUICKPAY_SMS_CONTAINER = []
LAST_VOUCHER_NO = 0
TRACK_PRINT_LOOP = [0]
PRINT_TIME_SCORE = 0
PRINT_RESET_THRESHOLD = 0
SERVER_NOT_RESPONDING = False
SNR_CAPTION_TIME = 0
TEMP_URL = ""
GHOSTSCRIPT_PATH = 'C:/Program Files (x86)/gs/gs9.55.0/bin/gswin32c.exe'
GSPRINT_PATH = ""
DOCUMENT = ""
SMS = "0"
START = False

COUNT = 0
COLOR = "green"


def kuickpay_sms():
    if not kuickpay_queue.empty():
        value = kuickpay_queue.get()
        name = value[1]
        due_date = value[6]
        amount = value[5]
        voucher_no = value[2]
        cell_number = value[4]
        date_obj = datetime.strptime(str(due_date).split(" ")[0], "%Y-%m-%d")
        kuickpay_formatted_date = date_obj.strftime("%d-%b-%Y")

        baseURL = "https://api.kuickpay.com"
        url = f"{baseURL}/api/sendSMS"

        headers = {
            "Content-Type": "application/json",
            "username": "INDUNIADMIN",
            "password": "Admin@12345"
        }

        data = {
             "consumerDetail": name,
              "merchantName": "Indus Universty",
              "dueDate": kuickpay_formatted_date,
              "amount": amount,
              "consumerNo": voucher_no,
              "cellNumberSMS": cell_number
        }

        try:
            try:
                response = requests.post(url, headers=headers, json=data, proxies=proxies)
            except requests.exceptions.ProxyError:
                response = requests.post(url, headers=headers, json=data)
            response.raise_for_status()
            print("Response:", response.text)
        except requests.exceptions.HTTPError as errh:
            print(f"HTTP Error: {errh}")
        except requests.exceptions.RequestException as err:
            print(f"Error: {err}")

    root.after(100, kuickpay_sms)





def download(url: str, dest_folder: str, id: str, timeout=60, max_retries=1):
    if not os.path.exists(dest_folder):
        os.makedirs(dest_folder)  # create folder if it does not exist

    filename = id  # file name based on ID
    file_path = os.path.join(dest_folder, filename)

    attempts = 0
    for i in range(max_retries):
        try:
            with urllib.request.urlopen(url, timeout=timeout) as response, open(file_path, 'wb') as out_file:
                data = response.read()  # Read the response data
                out_file.write(data)  # Write the data to a file
            return file_path
        except Exception as e:
            if attempts == max_retries:
                print(f"Server Not Responding... Failed to download {url}")
                ALREADY_PRINTED_VOUCHER.pop()
                return "Server Not Responding"
            else:
                print(f"Retrying download ...")
                attempts += 1


def search(x):
    input_queue.put(punch_rfid.get())
    punch_rfid.delete(0, 'end')


def find_and_print(rfid_code):
    global COUNT
    global COLOR
    global SMS
    global LAST_VOUCHER_NO
    global GHOSTSCRIPT_PATH, GSPRINT_PATH
    global SNR_CAPTION_TIME, SERVER_NOT_RESPONDING
    current_date = datetime.now()
    today_date = current_date.strftime("%d-%b-%y")

    with open("ghostscript_path.ini", "r") as file:
        for line in file:
            value = line.strip()
            if value.__contains__("GSPRINT_PATH = "):
                GSPRINT_PATH = value.split("GSPRINT_PATH = ")[1]
                if value.__contains__("SMS = "):
                    SMS = value.split("SMS = ")[1]

    rfids = get_pending_vouchers(rfid_code)
    if rfids == "Server Not Responding":
        SERVER_NOT_RESPONDING = True
        SNR_CAPTION_TIME = time.time() + 5
        rfids = None

    rfid = ""
    if rfids:
        for stu_data in rfids:
            voucher_no = str(stu_data[2])
            if voucher_no not in ALREADY_PRINTED_VOUCHER:
                rfid = stu_data
                break
        else:
            rfid = rfids[0]

    if rfid:
        name = rfid[1]
        voucher_no = str(rfid[2])
        if voucher_no not in ALREADY_PRINTED_VOUCHER or checkbox_var.get():
            if rfid[4] is not None:
                if SMS == "1":
                    kuickpay_queue.put(rfid)
            LAST_VOUCHER_NO = voucher_no
            ALREADY_PRINTED_VOUCHER.append(voucher_no)
            with open("voucher_history/" + today_date + ".txt", "a") as file:
                # Append a new line with a value
                new_value = str(voucher_no)
                file.write(new_value + "\n")
            COLOR = "yellow"
            status_label.configure(text=f'Processing Voucher... {name}... Please Wait')
            status_label_placement()
            narration = "Voucher Printed Successfully"
            narration_tag = "found"
            url = TEMP_URL.split('student_id=')[0] + "student_id=" + rfid[0] + "&v_voucher_no=" + str(rfid[2])

            file_name = str(rfid[2]) + '.pdf'
            file_location = "b/"

            if not os.path.exists(file_location + file_name):
                file_download = download(url, file_location, file_name)
                if file_download == "Server Not Responding":
                    narration = "Server Not Responding"
                    narration_tag = "not_found"
                    SERVER_NOT_RESPONDING = True
                    SNR_CAPTION_TIME = time.time() + 5
                    student_id, voucher_no, status, tag = rfid[0], rfid[2], narration, narration_tag
                    update_log(COUNT, student_id, name, voucher_no, status, tag)
                    COUNT += 1

            if os.path.exists(file_location + file_name):
                currentprinter = win32print.GetDefaultPrinter()
                ghostscript_cmd = (
                    f'-ghostscript "{GHOSTSCRIPT_PATH}" '
                    f'-printer "{currentprinter}" '
                    f'-dFirstPage=1 -dLastPage=1 '  # Specify the page range to print only the first page
                    f'-dOrientation=1 '  # Keep the orientation option
                    f'"{os.path.abspath(file_location + file_name)}"'  # The file to print
                )
                win32api.ShellExecute(0, 'open', GSPRINT_PATH, ghostscript_cmd, '.', 0)
                student_id, voucher_no, status, tag = rfid[0], rfid[2], narration, narration_tag
                update_log(COUNT, student_id, name, voucher_no, status, tag)
                COUNT += 1
                print("Voucher Print Successfully")
                checkbox_var.set(False)

        else:
            if LAST_VOUCHER_NO != voucher_no:
                COLOR = "red"
                student_id, voucher_no, status, tag = rfid[0], rfid[2], "Voucher Already Printed", "not_found"
                update_log(COUNT, student_id, name, voucher_no, status, tag)
                COUNT += 1
                status_label.configure(text=f'{name}, {voucher_no} Voucher Already Printed...')
                checkbox_var.set(False)
                LAST_VOUCHER_NO = str(voucher_no)
    else:
        COLOR = "red"
        status_label.configure(text=f'No Voucher Found')
        update_log(COUNT, "Unknown", "Unknown", "Unknown", "No Voucher Found", "not_found")
        COUNT += 1
        checkbox_var.set(False)


def print_job_checker():
    # print(len(TRACK_PRINT_LOOP))
    global DOCUMENT
    global PRINT_TIME_SCORE
    default_printer = win32print.GetDefaultPrinter()
    jobs = [1]  # Initialize with a dummy job to enter the loop
    while jobs:
        jobs = []
        phandle = win32print.OpenPrinter(default_printer)
        try:
            print_jobs = win32print.EnumJobs(phandle, 0, -1, 1)
            if print_jobs:
                pass
            else:
                if PRINT_TIME_SCORE != len(TRACK_PRINT_LOOP):
                    PRINT_TIME_SCORE = len(TRACK_PRINT_LOOP)
                    if PRINT_TIME_SCORE <= PRINT_RESET_THRESHOLD and PRINT_TIME_SCORE != 1:
                        if PRINT_TIME_SCORE != 0:
                            print(f"Retrying Print... {DOCUMENT}")
                            currentprinter = win32print.GetDefaultPrinter()
                            ghostscript_cmd = (
                                f'-ghostscript "{GHOSTSCRIPT_PATH}" '
                                f'-printer "{currentprinter}" '
                                f'-dFirstPage=1 -dLastPage=1 '  # Specify the page range to print only the first page
                                f'-dOrientation=1 '  # Keep the orientation option
                                f'"{os.path.abspath(DOCUMENT)}"'  # The file to print
                            )
                            win32api.ShellExecute(0, 'open', GSPRINT_PATH, ghostscript_cmd, '.', 0)
            try:
                if print_jobs:
                    jobs.extend(list(print_jobs))
                    for job in print_jobs:
                        if DOCUMENT != job.get("pDocument", "Unknown Document"):
                            DOCUMENT = job.get("pDocument", "Unknown Document")
                            TRACK_PRINT_LOOP.clear()
                            TRACK_PRINT_LOOP.append(DOCUMENT)
                        else:
                            TRACK_PRINT_LOOP.append(DOCUMENT)

            except Exception as e:
                error_message = str(e)
                print(f"Failed to print the document: {DOCUMENT}, Error: {error_message}")
                if ALREADY_PRINTED_VOUCHER:
                    ALREADY_PRINTED_VOUCHER.pop()
                return error_message
            finally:
                win32print.ClosePrinter(phandle)
        except Exception as e:
            print("Unhandled Exception")

        if not jobs:
            return None
        else:
            return "Printing Voucher... Please Wait !!!"


def fetching_data():
    global COUNT
    current_date = datetime.now()
    today_date = current_date.strftime("%d-%b-%y")

    if not os.path.exists("voucher_history/" + today_date + ".txt"):
        with open("voucher_history/" + today_date + ".txt", "w") as file:
            pass
        existing_pdf_files = [file for file in os.listdir('b') if file.endswith(".pdf")]
        for file in existing_pdf_files:
            os.remove('b/' + file)

    if os.path.exists("voucher_history/" + today_date + ".txt"):
        with open("voucher_history/" + today_date + ".txt", "r") as file:
            for line in file:
                COUNT += 1
                ALREADY_PRINTED_VOUCHER.append(line.strip())

    today_report_url = ""
    threshold = 0
    with open("ghostscript_path.ini", "r") as file:
        for line in file:
            value = line.strip()
            if value.__contains__("URL_PATH = "):
                today_report_url = value.split("URL_PATH = ")[1]
            if value.__contains__("PRINT_RESET_THRESHOLD = "):
                threshold = int(value.split("PRINT_RESET_THRESHOLD = ")[1])

    global TEMP_URL
    global PRINT_RESET_THRESHOLD
    TEMP_URL = today_report_url
    PRINT_RESET_THRESHOLD = threshold


def blink_label():
    global is_visible
    if is_visible:
        # status_label.configure(text_color=root.cget('background'))
        status_label.configure(text_color='white')
    else:
        status_label.configure(text_color=COLOR)
    is_visible = not is_visible
    root.after(500, blink_label)


def status_label_placement():
    root.update_idletasks()
    label_width = status_label.winfo_width()
    if label_width > WIDTH:
        font = 50
        F_new = font * (WIDTH / label_width)
        status_label.configure(font=("Roboto", int(F_new) - 8, 'bold'))
        label_width = status_label.winfo_width()
    status_label.place(x=(int(WIDTH / 2) - int(label_width / 2)), y=250)


def update_log(COUNT, student_id, student_name, voucher_no, status, tag):
    file_searched_log.insert(parent="", index='end', iid=COUNT, text=COUNT + 1,
                             values=(
                                 student_id, student_name, str(voucher_no),
                                 status), tags=(tag,))


def process_worker():
    global COLOR
    global SERVER_NOT_RESPONDING
    TIME = 0
    while True:
        status_label_placement()
        if not input_queue.empty():
            value = input_queue.get()
            find_and_print(value)
            file_searched_log.yview_moveto(1)
            root.update()
        if print_job_checker():
            COLOR = 'yellow'
            status_label.configure(text=print_job_checker())
            TIME = time.time() + 2
        else:
            if SERVER_NOT_RESPONDING:
                if SNR_CAPTION_TIME > time.time():
                    COLOR = 'red'
                    status_label.configure(text="SERVER NOT RESPONDING")
                else:
                    COLOR = 'green'
                    SERVER_NOT_RESPONDING = FALSE
            if time.time() > TIME and not SERVER_NOT_RESPONDING:
                COLOR = 'green'
                status_label.configure(text="SCAN YOUR RFID TO GET VOUCHER")


customtkinter.set_appearance_mode("dark")
customtkinter.set_default_color_theme("blue")
WIDTH, HEIGHT = 1200, 800

# Front End
root = customtkinter.CTk()
root.title("Automated Voucher Printing System")
root.geometry(f"{WIDTH}x{HEIGHT}")
root.resizable(0, 0)

checkbox_var = BooleanVar()

frame = customtkinter.CTkFrame(master=root, height=390)
frame.pack(pady=10, padx=10, fill="both", expand=False)

frame2 = customtkinter.CTkFrame(master=root, height=370)
frame2.pack(pady=10, padx=10, fill="both", expand=False)

image_path = "iu.png"
image = customtkinter.CTkImage(light_image=Image.open(image_path),
                               dark_image=Image.open(image_path),
                               size=(200, 100))

is_visible = True

# Create a label to display the image
iu_logo = customtkinter.CTkLabel(frame, image=image, text="")
iu_logo.place(x=WIDTH - 250, y=0)

application_name = customtkinter.CTkLabel(master=frame, text="Automated Voucher Printing", font=("Roboto", 40, 'bold'))
application_name.place(x=40, y=40)

punch_rfid = customtkinter.CTkEntry(master=frame, placeholder_text="Scan RFID", width=500, font=("Roboto", 50))
punch_rfid.place(x=350, y=150)
punch_rfid.focus_set()

punch_rfid.bind("<Return>", search)

status_label = customtkinter.CTkLabel(master=frame, text="SCAN YOUR RFID TO GET VOUCHER", font=("Roboto", 50, 'bold'))
status_label.place(x=400, y=270)

checkbox = customtkinter.CTkCheckBox(master=frame, text="Force Print", variable=checkbox_var, onvalue=True, offvalue=False)
checkbox.place(x=550, y=335)

style = ttk.Style()
style.theme_use("default")
style.configure("Treeview",
                background="#2B2B2B",
                foreground="#7F7F7F",
                fieldbackground="#2B2B2B",
                rowheight=35,
                )
style.configure('Treeview.Heading', background='black', foreground='#2FA572')
style.map('Treeview', background=[('selected', 'light green')], foreground=[('selected', '#7F7F7F')])

file_searched_log = ttk.Treeview(frame2, height=11, selectmode='none')
# file_searched_log.tag_configure('found', foreground='blue', background="#7F7F7F")
file_searched_log.tag_configure('found', foreground='white', background="#2B2B2B", font=('Arial', 12))
file_searched_log.tag_configure('not_found', foreground='red', background="black", font=('Arial', 12))
file_searched_log['columns'] = ("Student ID", "Student Name", "Voucher No", "Status")
file_searched_log.column("#0", minwidth=10, width=55)
file_searched_log.column("Student ID", width=150)
file_searched_log.column("Student Name", width=400)
file_searched_log.column("Voucher No", width=150)
file_searched_log.column("Status", width=400)

file_searched_log.heading("#0", text="S. No")
file_searched_log.heading("Student ID", text="Student ID")
file_searched_log.heading("Student Name", text="Student Name")
file_searched_log.heading("Voucher No", text="Voucher No")
file_searched_log.heading("Status", text="Status")

file_searched_log.pack(side='left', fill='y')

scrollbar = Scrollbar(frame2, orient="vertical", command=file_searched_log.yview)
scrollbar.pack(side="right", fill="y")
file_searched_log.configure(yscrollcommand=scrollbar.set)



fetching_data()
blink_label()
process_thread = threading.Thread(target=process_worker, daemon=True)
process_thread.start()
sms = '0'
with open("ghostscript_path.ini", "r") as file:
    for line in file:
        value = line.strip()
        if value.__contains__("SMS = "):
            sms = value.split("SMS = ")[1]
if sms == '1':
    kuickpay_sms_thread = threading.Thread(target=kuickpay_sms, daemon=True)
    kuickpay_sms_thread.start()
root.mainloop()
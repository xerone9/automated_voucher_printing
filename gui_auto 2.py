import customtkinter
from tkinter import *
from tkinter import ttk
from PIL import Image
import time

import os
import requests
import urllib.request

import threading
import queue
from datetime import datetime

import win32api
import win32print
import win32timezone


input_queue = queue.Queue()

VOUCHER_CONTAINER = {}
ALREADY_PRINTED_RFID = []
TEMP_URL = ""
GHOSTSCRIPT_PATH = 'C:/Program Files (x86)/gs/gs9.55.0/bin/gswin32c.exe'
GSPRINT_PATH = ""

COUNT = 0
COLOR = "green"


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


def search(x):
    input_queue.put(punch_rfid.get())
    punch_rfid.delete(0, 'end')


def find_and_print(rfid):
    global COUNT
    global COLOR
    global GHOSTSCRIPT_PATH, GSPRINT_PATH
    current_date = datetime.now()
    today_date = current_date.strftime("%d-%b-%y")

    with open("ghostscript_path.ini", "r") as file:
        for line in file:
            value = line.strip()
            if value.__contains__("GSPRINT_PATH = "):
                GSPRINT_PATH = value.split("GSPRINT_PATH = ")[1]

    if rfid in VOUCHER_CONTAINER:
        name = VOUCHER_CONTAINER[rfid][1]
        if rfid not in ALREADY_PRINTED_RFID:
            ALREADY_PRINTED_RFID.append(rfid)
            with open(today_date + ".txt", "a") as file:
                # Append a new line with a value
                new_value = rfid
                file.write(new_value + "\n")
            COLOR = "yellow"
            status_label.configure(text=f'Processing Voucher... {name}... Please Wait')
            status_label_placement()
            student_id, voucher_no, status, tag = VOUCHER_CONTAINER[rfid][0], VOUCHER_CONTAINER[rfid][2], "Voucher Printed Successfully", "found"
            update_log(COUNT, student_id, name, voucher_no, status, tag)
            COUNT += 1
            url = TEMP_URL.split('student_id=')[0] + "student_id=" + VOUCHER_CONTAINER[rfid][0] + "&v_voucher_no=" + \
                  VOUCHER_CONTAINER[rfid][2]

            file_name = VOUCHER_CONTAINER[rfid][1] + '.pdf'
            file_location = "b/"
            try:
                download(url, file_location, file_name)
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
            COLOR = "red"
            student_id, voucher_no, status, tag = VOUCHER_CONTAINER[rfid][0], \
                                                  VOUCHER_CONTAINER[rfid][
                                                      2], "Voucher Already Printed", "not_found"
            update_log(COUNT, student_id, name, voucher_no, status, tag)
            COUNT += 1
            status_label.configure(text=f'{name}, {voucher_no} Voucher Already Printed...')
    else:
        COLOR = "red"
        status_label.configure(text=f'No Voucher Found')
        update_log(COUNT, "Unknown", "Unknown", "Unknown", "No Voucher Found", "not_found")
        COUNT += 1


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
            # return str(document).split(".pdf")[0] + " Printing Voucher..."


def fetching_data():
    current_date = datetime.now()
    today_date = current_date.strftime("%d-%b-%y")

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
                voucher_container[str(i['student_id_barcode'])] = [i['student_id'], i['student_name'],
                                                                   str(i['student_voucher'])]

    global TEMP_URL, VOUCHER_CONTAINER
    TEMP_URL = today_report_url
    VOUCHER_CONTAINER = voucher_container


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
    status_label.place(x=(int(WIDTH / 2) - int(label_width / 2)), y=250)


def update_log(COUNT, student_id, student_name, voucher_no, status, tag):
    file_searched_log.insert(parent="", index='end', iid=COUNT, text=COUNT + 1,
                             values=(
                                 student_id, student_name, str(voucher_no),
                                 status), tags=(tag,))


def process_worker():
    global COLOR
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
            if time.time() > TIME:
                COLOR = 'green'
                status_label.configure(text="SCAN YOUR RFID TO GET VOUCHER")


customtkinter.set_appearance_mode("dark")
customtkinter.set_default_color_theme("blue")
WIDTH, HEIGHT = 1200, 800

#Front End
root = customtkinter.CTk()
root.title("Automated Voucher Printing System")
root.geometry(f"{WIDTH}x{HEIGHT}")
root.resizable(0, 0)

frame = customtkinter.CTkFrame(master=root, height=350)
frame.pack(pady=10, padx=10, fill="both", expand=False)

frame2 = customtkinter.CTkFrame(master=root, height=410)
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

status_label = customtkinter.CTkLabel(master=frame, text="SCAN YOUR RFID TO GET VOUCHER", font=("Roboto", 20, 'bold'))
status_label.place(x=400, y=250)

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
process_thread = threading.Thread(target=process_worker)
process_thread.start()
root.mainloop()
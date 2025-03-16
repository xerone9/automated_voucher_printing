import customtkinter
from tkinter import *
from tkinter import ttk
from PIL import Image
import time

import os
import urllib.request
import urllib.error

import threading
import queue
from datetime import datetime

import win32api
import win32print
import win32timezone

from get_voucher_from_database import get_pending_vouchers

input_queue = queue.Queue()

ALREADY_PRINTED_VOUCHER = []
TEMP_URL = ""
GHOSTSCRIPT_PATH = 'C:/Program Files (x86)/gs/gs9.55.0/bin/gswin32c.exe'
GSPRINT_PATH = ""

COUNT = 0
COLOR = "green"


def download(url: str, dest_folder: str, id: str, timeout=60, max_retries=3):
    if not os.path.exists(dest_folder):
        os.makedirs(dest_folder)  # create folder if it does not exist

    filename = id  # file name based on ID
    file_path = os.path.join(dest_folder, filename)

    try:
        with urllib.request.urlopen(url, timeout=timeout) as response, open(file_path, 'wb') as out_file:
            data = response.read()  # Read the response data
            out_file.write(data)  # Write the data to a file
        return file_path

    except Exception as e:
        print(f"Server Not Responding... Failed to download {url}")
        ALREADY_PRINTED_VOUCHER.pop()
        return None


def search(x):
    input_queue.put(punch_rfid.get())
    punch_rfid.delete(0, 'end')


def find_and_print(rfid_code):
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

    rfid = get_pending_vouchers(rfid_code)

    if rfid:
        name = rfid[1]
        voucher_no = str(rfid[2])
        if voucher_no not in ALREADY_PRINTED_VOUCHER or checkbox_var.get():
            ALREADY_PRINTED_VOUCHER.append(voucher_no)
            with open("voucher_history/" + today_date + ".txt", "a") as file:
                # Append a new line with a value
                new_value = str(voucher_no)
                file.write(new_value + "\n")
            COLOR = "yellow"
            status_label.configure(text=f'Processing Voucher... {name}... Please Wait')
            status_label_placement()
            student_id, voucher_no, status, tag = rfid[0], rfid[2], "Voucher Printed Successfully", "found"
            update_log(COUNT, student_id, name, voucher_no, status, tag)
            COUNT += 1
            url = TEMP_URL.split('student_id=')[0] + "student_id=" + rfid[0] + "&v_voucher_no=" + str(rfid[2])

            file_name = str(rfid[2]) + '.pdf'
            file_location = "b/"

            if not os.path.exists(file_location + file_name):
                download(url, file_location, file_name)

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
                print("Ghostscript Error ---> " + str(print_job_checker()))
                print("Voucher Print Successfully\n")
                checkbox_var.set(False)

        else:
            COLOR = "red"
            student_id, voucher_no, status, tag = rfid[0], rfid[2], "Voucher Already Printed", "not_found"
            update_log(COUNT, student_id, name, voucher_no, status, tag)
            COUNT += 1
            status_label.configure(text=f'{name}, {voucher_no} Voucher Already Printed...')
            checkbox_var.set(False)
    else:
        COLOR = "red"
        status_label.configure(text=f'No Voucher Found')
        update_log(COUNT, "Unknown", "Unknown", "Unknown", "No Voucher Found", "not_found")
        COUNT += 1
        checkbox_var.set(False)


def print_job_checker():
    default_printer = win32print.GetDefaultPrinter()
    jobs = [1]  # Initialize with a dummy job to enter the loop
    document = "No Document"
    retry_count = 0

    while jobs:
        jobs = []
        phandle = win32print.OpenPrinter(default_printer)

        try:
            print_jobs = win32print.EnumJobs(phandle, 0, -1, 1)
            if print_jobs:
                jobs.extend(list(print_jobs))
            for job in print_jobs:
                document = job["pDocument"]
        except Exception as e:
            error_message = str(e)
            print(error_message)
            ALREADY_PRINTED_VOUCHER.pop()
            return error_message
        finally:
            win32print.ClosePrinter(phandle)
        if not jobs:
            return None
        else:
            return "Printing Voucher... Please Wait !!!"


def fetching_data():
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
                ALREADY_PRINTED_VOUCHER.append(line.strip())

    today_report_url = ""
    with open("ghostscript_path.ini", "r") as file:
        for line in file:
            value = line.strip()
            if value.__contains__("URL_PATH = "):
                today_report_url = value.split("URL_PATH = ")[1]

    global TEMP_URL
    TEMP_URL = today_report_url


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
        status_label.configure(font=("Roboto", int(F_new), 'bold'))
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
process_thread = threading.Thread(target=process_worker)
process_thread.start()
root.mainloop()
import win32print
import time


def print_job_checker():
    """
    Prints out all jobs in the print queue every 5 seconds for the default printer.
    """
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
            return document + " Printing Voucher..."


if __name__ == "__main__":
    while True:
        if print_job_checker():
            print(print_job_checker())

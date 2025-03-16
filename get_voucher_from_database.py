import os
import cx_Oracle
from datetime import datetime
from collections import Counter

IP = ""
os.environ['PATH'] = 'C:/oracle/instantclient_19_3;' + os.environ['PATH']
with open("ghostscript_path.ini", "r") as file:
    for line in file:
        value = line.strip()
        if value.__contains__("ip = "):
            IP = value.split("ip = ")[1]


def dual_sort(value):
    last_two_digits = value[2] % 100
    return (-last_two_digits, -value[2])


def get_pending_vouchers(id, max_retries = 3):
    attempts = 0
    for i in range(max_retries):
        try:
            sql_code = f"SELECT student_id, student_name, voucher_no, rfid, student_cell_no, amount, due_date FROM print_vouchers WHERE (rfid = '{id}' OR student_id = '{id}') ORDER BY voucher_no DESC"

            # print("LD_LIBRARY_PATH:", os.environ.get('LD_LIBRARY_PATH'))
            # print("PATH:", os.environ.get('PATH'))

            # Replace these with your actual values
            username = 'usmanaccounts'
            password = 'usman_21022024'
            ip = IP
            port = '1521'  # Default Oracle port is 1521
            connect_string = 'orcl'

            # Construct the connection string
            dsn = cx_Oracle.makedsn(ip, port, service_name=connect_string)

            # Establish the connection
            connection = cx_Oracle.connect(username, password, dsn)

            data = []
            # Create a cursor
            cursor = connection.cursor()
            cursor.execute(sql_code)
            result = cursor.fetchall()
            for row in result:
                if id == row[3]:
                    data.append(row)
                else:
                    if row[5] <= 10000:
                        data.append(row)

            cursor.close()
            connection.close()
            if len(data) > 0:
                sorted_data = sorted(data, key=dual_sort)
                return sorted_data
            else:
                return None
        except Exception as e:
            print(e)
            if attempts == max_retries:
                if str(e).__contains__("ORA-12545"):
                    return "Server Not Responding"
            else:
                attempts += 1
                print("Retrying Fetching Data...")



# Kuickpay Inegration and return all pending vouchers
def get_pending_vouchers_60(rfid):
    current_date = datetime.now()
    today_date = current_date.strftime("%d-%b-%y")

    ALREADY_PRINTED_VOUCHER = []
    with open("voucher_history/" + today_date + ".txt", "r") as file:
        for line in file:
            ALREADY_PRINTED_VOUCHER.append(line.strip())
    sql_code = f"SELECT student_id, student_name, voucher_no, rfid, student_cell_no, amount, due_date FROM print_vouchers WHERE rfid = {rfid} ORDER BY voucher_no"
    os.environ['PATH'] = 'C:/oracle/instantclient_19_3;' + os.environ['PATH']

    # print("LD_LIBRARY_PATH:", os.environ.get('LD_LIBRARY_PATH'))
    # print("PATH:", os.environ.get('PATH'))

    # Replace these with your actual values
    username = 'usmanaccounts'
    password = 'usman_21022024'
    ip = 'lms.induscms.com'
    port = '1521'  # Default Oracle port is 1521
    connect_string = 'orcl'

    os.environ['PATH'] = 'C:/oracle/instantclient_19_3;' + os.environ['PATH']

    # Construct the connection string
    dsn = cx_Oracle.makedsn(ip, port, service_name=connect_string)

    # Establish the connection
    connection = cx_Oracle.connect(username, password, dsn)

    data = []
    # Create a cursor
    cursor = connection.cursor()
    cursor.execute(sql_code)
    result = cursor.fetchall()
    for row in result:
        data.append(row)

    cursor.close()
    connection.close()

    student_voucher_data = []
    if len(data) > 0:
        for i in data:
            stu_name = i[1]
            stu_due_date = i[6]
            stu_amount = i[5]
            stu_voucher = "06750" + str(i[2])
            stu_cell_number = i[4]
            kuickpay_data = [stu_name, stu_due_date, stu_amount, stu_voucher, stu_cell_number]
            if str(i[2]) not in ALREADY_PRINTED_VOUCHER:
                each_voucher_data = [[i[0], i[1], i[2]]]
                each_voucher_data.append(kuickpay_data)
                each_voucher_data.append(True)
                student_voucher_data.append(each_voucher_data)
            else:
                each_voucher_data = [[i[0], i[1], i[2]]]
                each_voucher_data.append(kuickpay_data)
                each_voucher_data.append(False)
                student_voucher_data.append(each_voucher_data)
        else:
            return student_voucher_data
    else:
        return None



def get_pending_vouchers_old(rfid):
    sql_code = f"SELECT student_id, student_name, voucher_no, rfid, student_cell_no, amount, due_date FROM print_vouchers WHERE rfid = {rfid} ORDER BY voucher_no"
    os.environ['PATH'] = 'C:/oracle/instantclient_19_3;' + os.environ['PATH']

    # print("LD_LIBRARY_PATH:", os.environ.get('LD_LIBRARY_PATH'))
    # print("PATH:", os.environ.get('PATH'))

    # Replace these with your actual values
    username = 'usmanaccounts'
    password = 'usman_21022024'
    ip = 'lms.induscms.com'
    port = '1521'  # Default Oracle port is 1521
    connect_string = 'orcl'

    os.environ['PATH'] = 'C:/oracle/instantclient_19_3;' + os.environ['PATH']

    # Construct the connection string
    dsn = cx_Oracle.makedsn(ip, port, service_name=connect_string)

    # Establish the connection
    connection = cx_Oracle.connect(username, password, dsn)

    data = []
    # Create a cursor
    cursor = connection.cursor()
    cursor.execute(sql_code)
    result = cursor.fetchall()
    for row in result:
        data.append(row)

    cursor.close()
    connection.close()
    if len(data) > 0:
        return data[-1][:3]
    else:
        return None


def get_all_pending_vouchers(max_retries=3):
    for i in range(max_retries):
        try:
            sql_code = f"SELECT student_id, student_name, voucher_no, rfid, student_cell_no, amount, due_date FROM print_vouchers ORDER BY student_id"
            os.environ['PATH'] = 'C:/oracle/instantclient_19_3;' + os.environ['PATH']

            # print("LD_LIBRARY_PATH:", os.environ.get('LD_LIBRARY_PATH'))
            # print("PATH:", os.environ.get('PATH'))

            # Replace these with your actual values
            username = 'usmanaccounts'
            password = 'usman_21022024'
            ip = IP
            port = '1521'  # Default Oracle port is 1521
            connect_string = 'orcl'

            os.environ['PATH'] = 'C:/oracle/instantclient_19_3;' + os.environ['PATH']

            # Construct the connection string
            dsn = cx_Oracle.makedsn(ip, port, service_name=connect_string)

            # Establish the connection
            connection = cx_Oracle.connect(username, password, dsn)

            data = []
            # Create a cursor
            cursor = connection.cursor()
            cursor.execute(sql_code)
            result = cursor.fetchall()
            for row in result:
                data.append(row)

            cursor.close()
            connection.close()
            if len(data) > 0:
                return data
            else:
                print("hi")
                return None
        except Exception as e:
            print(e)
            print("Retrying Fetching Data.. ")


# data_1 = []
# if get_all_pending_vouchers():
#     count = 0
#     for i in get_all_pending_vouchers():
#         data_1.append(i[0])
#         count += 1
#         print(i)
#     print("\n" + str(count))


#
# counter = Counter(data_1)
#
# # Print items with count > 1
# for item, count in counter.items():
#     if count > 2:
#         print(item)
# print("======================\n")

# for i in get_pending_vouchers('960-2024'):
#     print(i)

if get_all_pending_vouchers():
    for i in get_all_pending_vouchers():
        print(i)

# IMPORTS
import datetime
import numpy as np
import sys
from tkinter import *
import os
import tkinter as tk
from tkinter import filedialog, messagebox, PhotoImage
import cx_Oracle
import pandas as pd
from tabulate import tabulate
from tkintertable import TableCanvas
from pandastable import Table, TableModel
import openpyxl
from sklearn.ensemble import IsolationForest

# GLOBAL VARIABLES

## FOR ADD_FILES LIST
file_list = []
dataframes = []

## FOR RETRIEVE FILES LIST
rfile_list = []
rdataframes = []

fromfile_list = []
tofile_list = []

## CONNECTION AND SCHEMA
global connection
global dirSchema

## USERNAME AND SCHEMA
global USER
USER = "SYSTEM"
global PASS
PASS = "supersonic20"
global PORT
PORT = "1521"
global SID
SID = "orcl"
global SCHEMA
SCHEMA = "C##NEWSCHEMA"


# ROOTS / WINDOWS

root = tk.Tk()
root.geometry("1500x820")
root.title("Main Interface")
root.withdraw()

# CONSTANTS

global leftFrame_B3
global label_A6


def check_connection(_user, _pass, _port, _SID):
    global connection
    username = _user
    password = _pass
    SID = _SID
    port = _port

    URL = username + "/" + password + "@localhost:" + port + "/" + SID
    print(URL)

    connection = cx_Oracle.connect(URL)


def main_page():
    global root
    global connection
    global dirSchema
    global label_A6
    global listbox_A3
    global USER
    global PASS
    global PORT
    global SID

    check_connection(USER, PASS, PORT, SID)

    def view_file(file, label):
        # Use global variables
        global selected_file
        global label_A6

        selected_file = file

        if isinstance(file, pd.DataFrame):
            label.config(state=tk.NORMAL)
            label.delete("1.0", tk.END)
            table = Table(label, dataframe=file, showtoolbar=False, showstatusbar=False, editable=False)
            table.show()
            label.config(state=tk.DISABLED)

            # Adjust the size of the table to fit within the label
            table.redraw()
            table_width = table.winfo_reqwidth()
            table_height = table.winfo_reqheight()
            label.config(width=table_width, height=table_height)
        else:
            try:
                df = pd.read_excel(file)
            except Exception as e:
                label.config(state=tk.NORMAL)
                label.delete("1.0", tk.END)
                label.insert(tk.END, f"Error: {e}")
                label.config(state=tk.DISABLED)
                return

            excel_filename = "converted_dataframe.xlsx"
            df.to_excel(excel_filename, index=False)

            df_from_excel = pd.read_excel(excel_filename)

            with pd.option_context('display.max_columns', None):
                label.config(state=tk.NORMAL)
                label.delete("1.0", tk.END)
                table = Table(label, dataframe=df_from_excel, showtoolbar=False, showstatusbar=False, editable=False)
                table.show()
                label.config(state=tk.DISABLED)

                table.redraw()
                table_width = table.winfo_reqwidth()
                table_height = table.winfo_reqheight()
                label.config(width=table_width, height=table_height)

    def switch_file(event):
        global selected_file
        global listbox_A3
        global label_A6
        index = listbox_A3.curselection()
        if index:
            selected_file = file_list[index[0]]
            view_file(selected_file, label_A6)

    def add_files():
        global file_list
        global dataframes

        files = filedialog.askopenfilenames(filetypes=[("Excel Files", "*.xlsx")])
        for file in files:
            file_name = os.path.basename(file)
            file_list.append(file_name)
            dataframes.append(pd.DataFrame(pd.read_excel(file)))
            view_button = Button(leftFrame_A3, text=f"View {file_name}", command=lambda f=file: view_file(f, label_A6),
                                 bg="lightgreen")
            view_button.pack(pady=5)
        print(file_list)

    def xlsx_to_SQL():
        print("in")
        global file_list
        global dataframes
        global connection
        global label_B3
        schema = SCHEMA

        global dirname
        global dirpass
        global dirport
        global dirSID
        global status_label

        print(schema)
        cursor = connection.cursor()
        num = 0



        # Define a function to convert different data types to strings
        def convert_to_string(value):
            if isinstance(value, datetime.time):
                return value.strftime('%H:%M:%S')
            elif pd.notnull(value):
                return str(value)
            else:
                return None

        for df in dataframes:
            df = df.applymap(convert_to_string)
            table_name = os.path.splitext(file_list[num])[0]  # Remove the ".xlsx" extension

            columns = []
            for column_name, dtype in zip(df.columns, df.dtypes):
                if dtype == 'object':
                    column_type = 'VARCHAR2(500)'  # Adjust the maximum length as needed
                elif dtype == 'int64':
                    column_type = 'NUMBER'
                elif dtype == 'float64':
                    column_type = 'FLOAT'
                elif dtype == 'datetime64':
                    column_type = 'DATE'
                else:
                    column_type = 'VARCHAR2(500)'
                columns.append(f"{column_name} {column_type}")

            create_table_query = f"CREATE TABLE {schema}.\"{table_name}\" ({', '.join(columns)})"

            # Execute the CREATE TABLE statement
            cursor.execute(create_table_query)

            # Convert the DataFrame to a list of tuples
            data = [tuple(row) for row in df.applymap(convert_to_string).values]

            insert_query = f"INSERT INTO {schema}.\"{table_name}\" VALUES ({','.join([':' + str(i + 1) for i in range(len(df.columns))])})"
            print(insert_query)
            cursor.executemany(insert_query, data)

            num += 1

        connection.commit()
        cursor.close()
        connection.close()

        success_label = Label(root, text="SUCCESS: FILES HAVE BEEN UPLOADED", fg="green")
        label_B3.config(text="SUCCESS: FILES HAVE BEEN UPLOADED")


    def SQL_to_xlsx():
        global dirSchema
        global rfile_list
        global rdataframes

        global dirname
        global dirpass
        global dirport
        global dirSID
        global status_label

        username = USER
        password = PASS
        port = PORT
        sid = SID
        schema = SCHEMA

        check_connection(username, password, port, sid)

        cursor = connection.cursor()
        cursor.execute(f"SELECT table_name FROM all_tables WHERE owner = '{schema.upper()}'")
        table_names = [row[0] for row in cursor.fetchall()]

        print(table_names)

        cursor.close()

        # Iterate over each table
        for table_name in table_names:
            # Construct the query to retrieve the table data
            query = f"SELECT * FROM {schema}.\"{table_name}\""
            cursor = connection.cursor()
            cursor.execute(query)
            rows = cursor.fetchall()
            column_names = [col[0] for col in cursor.description]

            # Create a pandas DataFrame from the fetched rows and column names
            df = pd.DataFrame(rows, columns=column_names)
            rdataframes.append(df)
            rfile_list.append(table_name)
            cursor.close()

        # Close the database connection
        connection.close()

        for file_name, df in zip(rfile_list, rdataframes):
            view_button = Button(rightFrame_A3, text=f"View {file_name}",
                                 command=lambda f=df: view_file(f, rlabel_A6),
                                 bg="lightgreen")
            view_button.pack(pady=5)




        print("EXCEL FILES SUCCESSFULLY STORED IN DATABASE")
        #capture_console_output()

    global label_B5
    global label_B3

    root = Tk()  # create root window
    root.title("Main Interface")
    root.config(bg="lightgray")
    root.geometry("1500x820")

    # Configure row weights to control quadrant proportions
    root.grid_rowconfigure(0, weight=6)  # Top quadrants (60%)
    root.grid_rowconfigure(1, weight=4)  # Bottom quadrants (40%)

    # Left top quadrant - leftFrame_A
    leftFrame_A = Frame(root, bg="skyblue")
    leftFrame_A.grid(row=0, column=0, padx=10, pady=5, sticky="nsew")  # sticky="nsew" to fill the entire space

    # Right top quadrant - rightFrame_A
    rightFrame_A = Frame(root, bg="skyblue")
    rightFrame_A.grid(row=0, column=1, padx=10, pady=5, sticky="nsew")  # sticky="nsew" to fill the entire space

    # Divide the top left quadrant into three rows and two columns
    leftFrame_A.grid_rowconfigure(0, weight=0)
    leftFrame_A.grid_rowconfigure(1, weight=0)
    leftFrame_A.grid_rowconfigure(2, weight=9)

    leftFrame_A.grid_columnconfigure(0, weight=0)
    leftFrame_A.grid_columnconfigure(1, weight=10)

    # Top section in leftFrame_A - leftFrame_A1 (10%)
    leftFrame_A1 = Frame(leftFrame_A, bg="skyblue")
    leftFrame_A1.grid(row=0, column=0, padx=5, pady=5, sticky="nsew")  # sticky="nsew" to fill the entire space
    label_A1 = Label(leftFrame_A1, text=" Add Files To Upload ", font=("Arial", 14), bg="skyblue", fg="black")
    label_A1.pack(fill="both", expand=True)

    # Middle section in leftFrame_A - leftFrame_A2 (10%)
    leftFrame_A2 = Frame(leftFrame_A, bg="black")
    leftFrame_A2.grid(row=1, column=0, padx=5, pady=5, sticky="nsew")  # sticky="nsew" to fill the entire space
    button_A2 = Button(leftFrame_A2, text="Add Files", font=("Helvetica", 14), bg="black", fg="white",
                       command=add_files)
    button_A2.pack(fill="both", expand=True)

    # Bottom section in leftFrame_A - leftFrame_A3 (80%)

    leftFrame_A3 = Frame(leftFrame_A, bg="lightgreen")
    leftFrame_A3.grid(row=2, column=0, padx=5, pady=5, sticky="nsew")  # sticky="nsew" to fill the entire space
    listbox_A3 = Listbox(leftFrame_A3)
    scrollbar_A3 = Scrollbar(leftFrame_A3, bg="black")
    listbox_A3.config(yscrollcommand=scrollbar_A3.set)
    scrollbar_A3.config(command=listbox_A3.yview)
    scrollbar_A3.pack(side=RIGHT, fill=Y)

    # Top section in leftFrame_A - leftFrame_A4 (10%)
    leftFrame_A4 = Frame(leftFrame_A, bg="green")
    leftFrame_A4.grid(row=0, column=1, padx=5, pady=5, sticky="nsew")  # sticky="nsew" to fill the entire space
    label_A4 = Label(leftFrame_A4, text="Display Imported Files", font=("Arial", 16), bg="skyblue", fg="black")
    label_A4.pack(fill="both", expand=True)

    # Middle section in leftFrame_A - leftFrame_A5 (10%)
    leftFrame_A5 = Frame(leftFrame_A, bg="darkgreen")
    leftFrame_A5.grid(row=1, column=1, padx=5, pady=5, sticky="nsew")  # sticky="nsew" to fill the entire space
    label_A5 = Label(leftFrame_A5, text="", font=("Arial", 16), bg="skyblue", fg="white")
    label_A5.pack(fill="both", expand=True)

    # Bottom section in leftFrame_A - leftFrame_A6 (80%)
    leftFrame_A6 = Frame(leftFrame_A, bg="lightgreen")
    leftFrame_A6.grid(row=2, column=1, padx=5, pady=5, sticky="nsew")  # sticky="nsew" to fill the entire space
    label_A6 = Text(leftFrame_A6, font=("Oswald", 12), bg="lightgreen", fg="white", height = 20, width = 5)
    label_A6.pack(fill="both", expand=True)
    lDisplay_scrollbar = tk.Scrollbar(leftFrame_A)
    label_A6.config(yscrollcommand=lDisplay_scrollbar.set)
    lDisplay_scrollbar.config(command=label_A6.yview)

    # Divide the top right quadrant into three rows and two columns
    rightFrame_A.grid_rowconfigure(0, weight=0)
    rightFrame_A.grid_rowconfigure(1, weight=0)
    rightFrame_A.grid_rowconfigure(2, weight=9)

    rightFrame_A.grid_columnconfigure(0, weight=0)
    rightFrame_A.grid_columnconfigure(1, weight=10)

    # Top section in rightFrame_A - rightFrame_A1 (10%)
    rightFrame_A1 = Frame(rightFrame_A, bg="skyblue")
    rightFrame_A1.grid(row=0, column=0, padx=5, pady=5, sticky="nsew")  # sticky="nsew" to fill the entire space
    label_A4 = Label(rightFrame_A1, text="Retrieve from Database", font=("Arial", 14), bg="skyblue", fg="black")
    label_A4.pack(fill="both", expand=True)

    # Middle section in rightFrame_A - rightFrame_A2 (10%)
    rightFrame_A2 = Frame(rightFrame_A, bg="darkblue")
    rightFrame_A2.grid(row=1, column=0, padx=5, pady=5, sticky="nsew")  # sticky="nsew" to fill the entire space
    label_A5 = Button(rightFrame_A2, text="Retrieve Files", font=("Arial", 14), bg="black", fg="white", command=SQL_to_xlsx)
    label_A5.pack(fill="both", expand=True)

    # Bottom section in rightFrame_A - rightFrame_A3 (80%)
    rightFrame_A3 = Frame(rightFrame_A, bg="lightgreen")
    rightFrame_A3.grid(row=2, column=0, padx=5, pady=5, sticky="nsew")  # sticky="nsew" to fill the entire space
    rlistbox_A3 = Listbox(rightFrame_A3)
    rscrollbar_A3 = Scrollbar(rightFrame_A3, bg="black")
    rlistbox_A3.config(yscrollcommand=rscrollbar_A3.set)
    rscrollbar_A3.config(command=rlistbox_A3.yview)
    rscrollbar_A3.pack(side=RIGHT, fill=Y, expand=False)

    # Top section in rightFrame_A - rightFrame_A1 (10%)
    rightFrame_A4 = Frame(rightFrame_A, bg="blue")
    rightFrame_A4.grid(row=0, column=1, padx=5, pady=5, sticky="nsew")  # sticky="nsew" to fill the entire space
    label_A4 = Label(rightFrame_A4, text="Display Retrieved Files", font=("Arial", 16), bg="skyblue", fg="black")
    label_A4.pack(fill="both", expand=True)

    # Middle section in rightFrame_A - rightFrame_A2 (10%)
    rightFrame_A5 = Frame(rightFrame_A, bg="darkblue")
    rightFrame_A5.grid(row=1, column=1, padx=5, pady=5, sticky="nsew")  # sticky="nsew" to fill the entire space
    label_A5 = Label(rightFrame_A5, text="", font=("Arial", 16), bg="skyblue", fg="white")
    label_A5.pack(fill="both", expand=True)

    # Bottom section in rightFrame_A - rightFrame_A3 (80%)
    rightFrame_A6 = Frame(rightFrame_A, bg="lightblue")
    rightFrame_A6.grid(row=2, column=1, padx=5, pady=5, sticky="nsew")  # sticky="nsew" to fill the entire space
    rlabel_A6 = Text(rightFrame_A6, font=("Oswald", 12), bg="lightgreen", fg="white", height = 20, width = 5)
    rlabel_A6.pack(fill="both", expand=True)
    rDisplay_scrollbar = tk.Scrollbar(rightFrame_A)
    label_A6.config(yscrollcommand=rDisplay_scrollbar.set)
    rDisplay_scrollbar.config(command=rlabel_A6.yview)

    # Left bottom quadrant - leftFrame_B
    leftFrame_B = Frame(root, bg="skyblue")
    leftFrame_B.grid(row=1, column=0, padx=10, pady=5, sticky="nsew")  # sticky="nsew" to fill the entire space

    # Divide the bottom left quadrant into three rows
    leftFrame_B.grid_rowconfigure(0, weight=1)
    leftFrame_B.grid_rowconfigure(1, weight=1)
    leftFrame_B.grid_rowconfigure(2, weight=8)

    leftFrame_B.grid_columnconfigure(0, weight=5)
    leftFrame_B.grid_columnconfigure(0, weight=5)

    # Top section in leftFrame_B - leftFrame_B1 (10%)
    leftFrame_B1 = Frame(leftFrame_B, bg="skyblue")
    leftFrame_B1.grid(row=0, column=0, padx=5, pady=5, sticky="nsew")  # sticky="nsew" to fill the entire space
    label_B1 = Button(leftFrame_B1, text="Find Discrepancies", font=("Arial", 16), bg="black", fg="green", command=find_discrepanies)
    label_B1.pack(fill="both", expand=True)

    # Middle section in leftFrame_B - leftFrame_B2 (10%)
    leftFrame_B2 = Frame(leftFrame_B, bg="darkred")
    leftFrame_B2.grid(row=1, column=0, padx=5, pady=5, sticky="nsew")  # sticky="nsew" to fill the entire space
    label_B2 = Label(leftFrame_B2, text="Console", font=("Arial", 16), bg="skyblue", fg="black")
    label_B2.pack(fill="both", expand=True)

    # Bottom section in leftFrame_B - leftFrame_B3 (80%)
    leftFrame_B3 = Frame(leftFrame_B, bg="pink")
    leftFrame_B3.grid(row=2, column=0, padx=5, pady=5, sticky="nsew")  # sticky="nsew" to fill the entire space
    label_B3 = Label(leftFrame_B3, text="", font=("Hack", 16), bg="black", fg="green")
    label_B3.pack(fill="both", expand=True)

    # Top section in leftFrame_B - leftFrame_B1 (10%)
    leftFrame_B4 = Frame(leftFrame_B, bg="skyblue")
    leftFrame_B4.grid(row=0, column=1, padx=5, pady=5, sticky="nsew")  # sticky="nsew" to fill the entire space
    label_B4 = Button(leftFrame_B4, text="Download Output", font=("Arial", 16), bg="black", fg="green", command=download_output)
    label_B4.pack(fill="both", expand=True)

    # Middle section in leftFrame_B - leftFrame_B2 (10%)
    leftFrame_B5 = Frame(leftFrame_B, bg="skyblue")
    leftFrame_B5.grid(row=1, column=1, padx=5, pady=5, sticky="nsew")  # sticky="nsew" to fill the entire space
    label_B5 = Button(leftFrame_B5, text="Upload to Database", font=("Arial", 14), bg="black", fg="green", command=xlsx_to_SQL)
    label_B5.pack(fill="both", expand=True)

    # Bottom section in leftFrame_B - leftFrame_B3 (80%)
    leftFrame_B6 = Frame(leftFrame_B, bg="pink")
    leftFrame_B6.grid(row=2, column=1, padx=5, pady=5, sticky="nsew")  # sticky="nsew" to fill the entire space
    label_B6 = Label(leftFrame_B6, text="                  ", font=("Hack", 30), bg="skyblue", fg="green")
    label_B6.pack(fill="both", expand=True)

    # Right bottom quadrant - rightFrame_B
    rightFrame_B = Frame(root, bg="skyblue")
    rightFrame_B.grid(row=1, column=1, padx=10, pady=5, sticky="nsew")  # sticky="nsew" to fill the entire space

    # Divide the bottom right quadrant into three rows and two columns
    rightFrame_B.grid_rowconfigure(0, weight=1)
    rightFrame_B.grid_rowconfigure(1, weight=9)
    rightFrame_B.grid_rowconfigure(2, weight=0)

    rightFrame_B.grid_columnconfigure(0, weight=10)

    # Top section in rightFrame_B - rightFrame_B1 (10%)
    rightFrame_B1 = Frame(rightFrame_B, bg="skyblue")
    rightFrame_B1.grid(row=0, column=0, padx=5, pady=5, sticky="nsew")  # sticky="nsew" to fill the entire space
    label_B4 = Label(rightFrame_B1, text="Results", font=("Arial", 16), bg="skyblue", fg="black")
    label_B4.pack(fill="both", expand=True)

    # Middle section in rightFrame_B - rightFrame_B2 (10%)
    rightFrame_B2 = Frame(rightFrame_B, bg="skyblue")
    rightFrame_B2.grid(row=1, column=0, padx=5, pady=5, sticky="nsew")  # sticky="nsew" to fill the entire space
    label_B5 = Label(rightFrame_B2, text="", font=("Arial", 8), bg="black", fg="green")
    label_B5.pack(fill="both", expand=True)

    # Configure column weights to fill the entire width of the root
    root.grid_columnconfigure(0, weight=1)
    root.grid_columnconfigure(1, weight=1)

    root.mainloop()


def find_discrepanies():
    global label_B3
    global dataframes_combine
    global fromfile_list
    global tofile_list

    dataframes_combine = []

    string = ""
    for i, file1 in enumerate(rfile_list):
        for j, file2 in enumerate(rfile_list):
            if file1 == "PurchaseOrder" and file2 == "GRNHeader":
                if rdataframes[i].loc[:, "PURCHASEORDERID"].equals(rdataframes[j].loc[:, "PURCHASEORDERID"]):
                    print("PURCHASE ORDER ID MATCHES")
                    string += "PURCHASE ORDER ID MATCHES\n"

                    fromfile_list.append("PurchaseOrder")
                    tofile_list.append("GRNHeader")
                else:
                    print("DISCREPANCY DECTECTED: PURCHASE ORDER ID")
                    string += "DISCREPANCY DECTECTED: INSIDE \"PurchaseOrder.xlsx\" and \"GRNHeader.xlsx\"\n"

                    df1_i = rdataframes[i]["PURCHASEORDERID"]
                    df1_j = rdataframes[j]["PURCHASEORDERID"]
                    df1 = pd.DataFrame({'PurchaseOrderId_PurchaseOrder': df1_i, 'PurchaseOrderId_GRNHeader': df1_j})

                    dataframes_combine.append(df1)
                    fromfile_list.append("PurchaseOrder")
                    tofile_list.append("GRNHeader")

            elif file1 == "VendorPayment" and file2 == "VendorPayment":
                if rdataframes[i].loc[:, "PAYMENTAMOUNT"].equals(rdataframes[j].loc[:, "SETTLEDAMOUNT"]):
                    print("PAYMENT MATCHES")
                    string += "\nPAYMENT MATCHES\n"

                    fromfile_list.append("VendorPayment")
                    tofile_list.append("VendorPayment")

                else:
                    print("DISCREPANCY DECTECTED: PAYMENT")
                    string += "DISCREPANCY DECTECTED: \"VendorPayment.xlsx\" and \"VendorPayment.xlsx\"\n"


                    df2_i = rdataframes[i]["PAYMENTAMOUNT"]
                    df2_j = rdataframes[j]["SETTLEDAMOUNT"]
                    non_matching_values_a = df2_i[~df2_i.isin(df2_j)]

                    df = pd.concat([df2_i, df2_j], axis=1)

                    df_z = df.loc[non_matching_values_a.index, :]

                    dataframes_combine.append(df_z)
                    fromfile_list.append("VendorPayment")
                    tofile_list.append("VendorPayment")

            elif file1 == "Expensesheet" and file2 == "Expensesheet":
                print(rdataframes[i].columns)
                if (rdataframes[i]["EMPLOYEEID"] == rdataframes[j]["ACKNOWLEDGEBY"]).all():
                    print("EMPLOYEE ID MATCHES")
                    string += "EMPLOYEE ID MATCHES\n"

                    fromfile_list.append("Expensesheet")
                    tofile_list.append("Expensesheet")
                else:
                    print("DISCREPANCY DETECTED: EMPLOYEE ID")
                    string += "DISCREPANCY DETECTED: \"Expensesheet.xlsx\" and \"Expensesheet.xlsx\"\n"

                    df3_i = rdataframes[i]["EMPLOYEEID"]
                    df3_j = rdataframes[j]["ACKNOWLEDGEBY"]

                    df = pd.concat([df3_i, df3_j], axis=1)

                    # Create an empty list to store dictionaries of rows
                    rows_to_append = []

                    # Iterate through rows and compare values
                    for index, row in df.iterrows():
                        if row['EMPLOYEEID'] != row['ACKNOWLEDGEBY']:
                            rows_to_append.append(
                                {'employeeid': row['EMPLOYEEID'], 'acknowledgeby': row['ACKNOWLEDGEBY']})

                    # Create a new DataFrame from the list of dictionaries
                    df_z = pd.DataFrame(rows_to_append, columns=['employeeid', 'acknowledgeby'])

                    dataframes_combine.append(df_z)
                    fromfile_list.append("Expensesheet")
                    tofile_list.append("Expensesheet")

            elif file1 == "PurchaseOrder" and file2 == "PurchaseOrderDetail":
                if rdataframes[i]["PURCHASEORDERID"].isin(rdataframes[j]["PURCHASEORDERID"]).any():
                    print("PURCHASE ORDER ID EXISTS IN PURCHASE ORDER LIST")
                    string += "PURCHASE ORDER ID EXISTS IN PURCHASE ORDER LIST\n"

                    fromfile_list.append("PurchaseOrder")
                    tofile_list.append("PurchaseOrderDetail")
                else:
                    print("DISCREPANCY DETECTED: PURCHASE ORDER ID DOES NOT EXIST IN PURCHASE ORDER LIST")
                    string += "DISCREPANCY DETECTED: \"PurchaseOrderID.xlsx\" and \"PurchaseOrderDetail.xlsx\"\n"

                    rows_to_append = []

                    df4_i = rdataframes[i]["PURCHASEORDERID"]
                    x_row = df4_i.iloc[0]
                    df4_j = rdataframes[j]["PURCHASEORDERID"]

                    # Iterate through rows in DataFrame Y
                    for index, y_row in df4_j.iterrows():
                        if df4_i['PURCHASEORDER'] != y_row['PURCHASEORDERID']:
                            rows_to_append.append(
                                {'purchaseorderid1': x_row['PURCHASEORDERID'], 'purchaseorderid2': y_row['PURCHASEORDERID']})

                    # Create a new DataFrame from the list of dictionaries
                    df_z = pd.DataFrame(rows_to_append, columns=['purchaseorderid1', 'purchaseorderid2'])


                    dataframes_combine.append(df_z)
                    fromfile_list.append("PurchaseOrder")
                    tofile_list.append("PurchaseOrderDetail")


            elif file1 == "vendorPaymentDetails" and file2 == "VendorPayment":
                if set(rdataframes[i]["PAYMENTID"]).issubset(set(rdataframes[j]["PAYMENTID"])):
                    print("PAYMENT ID EXISTS IN PAYMENT ID LIST")
                    string += "PAYMENT ID EXISTS IN PAYMENT ID LIST"

                    fromfile_list.append("vendorPaymentDetails")
                    tofile_list.append("VendorPayment")
                else:
                    print("DISCREPANCY DETECTED: PAYMENT ID DOES NOT EXIST IN PAYMENT ID LIST")
                    string += "DISCREPANCY DETECTED: \"vendorPaymentDetails.xlsx\" and \"VendorPayment.xlsx\"\n"

                    df5_i = rdataframes[i]["PAYMENTID"]
                    df5_i = df5_i.to_frame()
                    df5_j = rdataframes[j]["PAYMENTID"]
                    df5_j = df5_j.to_frame()

                    paymentID_list = df5_j['PAYMENTID'].tolist()

                    # Create an empty list to store dictionaries of rows
                    rows_to_append = []

                    # Iterate through rows in df5_i
                    for index, df5_i_row in df5_i.iterrows():
                        paymentID = df5_i_row['PAYMENTID']

                        # Check if the paymentID exists in paymentID_list
                        if paymentID not in paymentID_list:
                            rows_to_append.append({'paymentID': paymentID})

                    # Create a new DataFrame from the list of dictionaries
                    df_z = pd.DataFrame(rows_to_append, columns=['PAYMENTID'])

                    dataframes_combine.append(df_z)
                    fromfile_list.append("vendorPaymentDetails")
                    tofile_list.append("VendorPayment")

    print("\n\n\n")
    for j, filename in enumerate(rfile_list):
        if filename == "GRNDetail":
            print("RECEIVED")
            print(rdataframes[j])
            data_df = rdataframes[j]
            data_df = data_df.dropna(axis=1)
            for column in data_df.columns:
                if data_df[column].dtype == object:  # Check if the column contains non-numeric (object) data
                    data_df = pd.get_dummies(data_df, columns=[column], drop_first=True)

            model = IsolationForest(contamination=0.05, random_state=42)
            model.fit(data_df)
            anomaly_scores = model.score_samples(data_df)
            for i, score in enumerate(anomaly_scores):
                if score <= -0.5:
                    string += f"Anomaly Score: {i + 1}: Strong Correlation for Anomaly (High Chance of Anomaly)\n"
                elif -0.5 < score < 0.5:
                    string += f"Anomaly Score: {i + 1}: Medium Correlation for Anomaly\n"
                elif score >= 0.5:
                    string += f"Anomaly Score: {i + 1}: Weak Correlation for Anomaly (Low Chance of Anomaly) \n"

            print("\n")
    label_B5.config(text=string)


def download_output():
    global dataframes_combine
    global fromfile_list
    global tofile_list

    writer = pd.ExcelWriter('output.xlsx', engine='openpyxl')

    for i, df in enumerate(dataframes_combine):
        column_names = df.columns.tolist()
        selected_df = df[column_names]
        sheet_name = f'Sheet{i + 1}'

        # Create a new DataFrame with the fromfile_list and tofile_list values
        header_df = pd.DataFrame([[fromfile_list[i], tofile_list[i]]], columns=["From File", "To File"])

        # Concatenate the header DataFrame with the selected_df
        final_df = pd.concat([header_df, selected_df], ignore_index=True)

        # Write the final_df to the Excel sheet
        final_df.to_excel(writer, sheet_name=sheet_name, index=False)

    # Save the Excel workbook
    writer.book.save(filename="output.xlsx")
    label_B3.config(text="SUCCESS: THE OUTPUT HAS BEEN DOWNLOADED")




# MAIN

main_page()
root.mainloop()


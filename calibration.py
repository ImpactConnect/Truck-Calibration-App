import os
import sqlite3
import tkinter as tk
from tkinter import messagebox
from tkinter import ttk

import pandas as pd
from tkcalendar import DateEntry

master = tk.Tk()
master.title("Root Window")
width = 1000
height = 650
screen_width = master.winfo_screenwidth()
screen_height = master.winfo_screenheight()
x = int((screen_width / 2) - (width / 2))
y = int((screen_height / 2) - (height / 2))
master.geometry(f"{width}x{height}+{x}+{y}")


def open_top_level():
    top_level = tk.Toplevel(master)
    CalibrationEntry(top_level)

def load_excel_file():
    file_path = os.path.expanduser("~/Documents/Calibration Record.xlsx")

    try:
        df = pd.read_excel(file_path)
    except:
        messagebox.showerror("Error", "Unable to access the file!")
        return

    # Clear the treeview before loading the new file
    treeview.delete(*treeview.get_children())

    # Set the treeview headings
    treeview['columns'] = list(df.columns)
    treeview['show'] = "headings"

    for col in treeview['columns']:
        treeview.heading(col, text=col)
        treeview.column(col, width=100, anchor='center', stretch=True, minwidth=100)  # Adjusted minwidth


    # Insert the data rows into the treeview
    df_rows = df.values.tolist()
    for row in df_rows:
        treeview.insert("", "end", values=row)

def view_record():
    selected_item = treeview.focus()
    if not selected_item:
        messagebox.showwarning("No Record Selected", "Please select a record to view.")
        return

    values = treeview.item(selected_item, 'values')
    if not values:
        return

    popup = tk.Toplevel(master)
    popup.title("Truck Calibration Details")

    # Create a frame for the top section (horizontal display)
    top_frame = ttk.Frame(popup)
    top_frame.pack(padx=10, pady=10)

    # Display the first six headings and values horizontally
    for i, heading in enumerate(HEADINGS[:6]):
        ttk.Label(top_frame, text=heading).grid(row=0, column=i, padx=5, pady=5)
        ttk.Label(top_frame, text=values[i]).grid(row=1, column=i, padx=5, pady=5)

    # Create a frame for the bottom section (vertical display)
    bottom_frame = ttk.Frame(popup)
    bottom_frame.pack(padx=10, pady=10)

    # Display the remaining headings and values vertically in four columns
    for i, heading in enumerate(HEADINGS[6:], start=6):
        row = (i - 6) % 7
        column = (i - 6) // 7

        ttk.Label(bottom_frame, text=heading).grid(row=row, column=column * 2, sticky='e', padx=5, pady=5)
        ttk.Label(bottom_frame, text=values[i]).grid(row=row, column=column * 2 + 1, sticky='w', padx=5, pady=5)

# def search_keyword():
#     keyword = search_entry.get().upper()
#
#     if not keyword:
#         messagebox.showwarning("Warning", "Please enter a keyword for search.")
#         return
#
#     file_path = os.path.expanduser("~/Documents/Calibration Record.xlsx")
#
#     try:
#         df = pd.read_excel(file_path)
#     except:
#         messagebox.showerror("Error", "Unable to access the file!")
#         return
#
#     found_rows = []
#
#     for _, row in df.iterrows():
#         truck_number = row['Truck Number']
#         if truck_number == keyword:
#             found_rows.append(row)
#
#     if not found_rows:
#         messagebox.showinfo("Search Results", "No matching records found.")
#         return
#
#     # Clear the treeview before displaying the search results
#     treeview.delete(*treeview.get_children())
#
#     # Set the treeview headings based on the DataFrame columns
#     treeview['columns'] = list(df.columns)
#     treeview['show'] = "headings"
#
#     for col in treeview['columns']:
#         treeview.heading(col, text=col)
#
#     # Insert the search results into the treeview
#     for _, row in pd.DataFrame(found_rows).iterrows():
#         values = row.values.tolist()
#         treeview.insert("", "end", values=values)

def search_keyword():
    keyword = search_entry.get().upper()

    if not keyword:
        messagebox.showwarning("Warning", "Please enter a keyword for search.")
        return

    selected_option = search_option.get()

    if selected_option == "Select Search Location":
        messagebox.showwarning("Warning", "Please select a search location from the drop-down box.")
        return

    if selected_option == "Search from Local":
        # Search from local file
        file_path = os.path.expanduser("~/Documents/Calibration Record.xlsx")
        try:
            df = pd.read_excel(file_path)
        except:
            messagebox.showerror("Error", "Unable to access the file!")
            return
    else:
        # Search from database (add your database search code here)
        # Replace the code below with your database search implementation
        messagebox.showinfo("Search Results", "Search from database not implemented yet.")
        return

    found_rows = []

    for _, row in df.iterrows():
        truck_number = row['Truck Number']
        if truck_number == keyword:
            found_rows.append(row)

    if not found_rows:
        messagebox.showinfo("Search Results", "No matching records found.")
        return

    # Clear the treeview before displaying the search results
    treeview.delete(*treeview.get_children())

    # Set the treeview headings based on the DataFrame columns
    treeview['columns'] = list(df.columns)
    treeview['show'] = "headings"

    for col in treeview['columns']:
        treeview.heading(col, text=col)
        treeview.column(col, width=100, anchor='center', stretch=True, minwidth=100)  # Adjusted minwidth


    # Insert the search results into the treeview
    for _, row in pd.DataFrame(found_rows).iterrows():
        values = row.values.tolist()
        treeview.insert("", "end", values=values)



class CalibrationEntry:
    def __init__(self, Calpage):
        self.Calpage = Calpage
        self.Calpage.title("Batch Analysis")
        width = 900
        height = 600
        screen_width = self.Calpage.winfo_screenwidth()
        screen_height = self.Calpage.winfo_screenheight()
        x = int((screen_width / 2) - (width / 2))
        y = int((screen_height / 2) - (height / 2))
        self.Calpage.geometry(f"{width}x{height}+{x}+{y}")


        self.parameters = ['Date', 'Truck Number', 'Calibrator', 'Transporter', 'Total Volume', 'Chassis Number',
                           'Comp1 Vol', 'Comp1 (NH)', 'Comp1 (OH)', 'Comp1 (Final UH)', 'Comp1 (First UH)',
                           'Comp1 (Final LH)', 'Comp1 (First LH)','Comp2 Vol', 'Comp2 (NH)', 'Comp2 (OH)', 'Comp2 (Final UH)', 'Comp2 (First UH)',
                           'Comp2 (Final LH)', 'Comp2 (First LH)', 'Comp3 Vol', 'Comp3 (NH)', 'Comp3 (OH)', 'Comp3 (Final UH)', 'Comp3 (First UH)',
                           'Comp3 (Final LH)', 'Comp3 (First LH)', 'Comp4 Vol', 'Comp4 (NH)', 'Comp4 (OH)', 'Comp4 (Final UH)', 'Comp4 (First UH)',
                           'Comp4 (Final LH)',  'Comp4 (First LH)',
                           'T-Table Initial', 'T-Table Final', 'Balloon Before', 'Balloon After', 'B-Spring No',
                           'Front Spring Number', 'Calibration Status'
                           ]

        self.value_entry = {}
        for entry in self.parameters:
            self.value_entry[entry] = ttk.Entry(self.Calpage)
        # FRAMES
        caseframe = ttk.Frame(self.Calpage)
        caseframe.pack()
        frame1 = tk.LabelFrame(caseframe, text='Product Details')
        frame1.grid(row=0, column=0, columnspan=2, sticky=tk.W)
        frame2 = tk.LabelFrame(caseframe, text='Comp 1')
        frame2.grid(row=1, column=0, columnspan=1, )
        frame3 = tk.LabelFrame(caseframe, text='Comp 2')
        frame3.grid(row=1, column=1, padx=10, pady=10)
        frame4 = tk.LabelFrame(caseframe, text='Comp 3')
        frame4.grid(row=2, column=0, columnspan=1, pady=10)
        frame5 = tk.LabelFrame(caseframe, text='Comp 4')
        frame5.grid(row=2, column=1, pady=10)
        frame6 = tk.LabelFrame(caseframe, text='Other Truck Info')
        frame6.grid(row=3, column=0, columnspan=2, pady=10)
        button_frame = ttk.Frame(caseframe)
        button_frame.grid(row=4, column=0, columnspan=2, pady=10)

        for i, entry in enumerate(self.parameters[:5]):
            label = ttk.Label(frame1, text=entry)
            label.grid(row=0, column=i * 2, pady=(10, 2), padx=(2, 5), sticky=tk.E)

            if entry == 'Date':
                self.value_entry[entry] = DateEntry(frame1, width=12, background='darkblue',
                                                    foreground='white', borderwidth=2, year=2023,
                                                    date_pattern='dd/MM/yyyy')
                self.value_entry[entry].grid(row=0, column=i * 2 + 1, pady=(10, 2), padx=(0, 10))

            elif entry == 'Calibrator':
                self.value_entry[entry] = ttk.Combobox(frame1, values=['Ola', 'Olu', 'Ade', 'Ife', 'Ire'], width=10)
                self.value_entry[entry].grid(row=0, column=i * 2 + 1, pady=(10, 2), padx=(0, 10))
            elif entry == 'Transporter':
                self.value_entry[entry] = ttk.Combobox(frame1, values=['Ola', 'Olu', 'Ade', 'Ife', 'Ire'], width=10)
                self.value_entry[entry].grid(row=0, column=i * 2 + 1, pady=(10, 2), padx=(0, 10))
            else:
                self.value_entry[entry] = ttk.Entry(frame1, width=10)
                self.value_entry[entry].grid(row=0, column=i * 2 + 1, pady=(10, 2), padx=(0, 10))


        # Create label and entry field for item[5]
        ttk.Label(frame1, text=self.parameters[5]).grid(row=1, column=2, pady=(2, 2), padx=(5, 5), )
        ttk.Entry(frame1, width=20).grid(row=1, column=3, pady=(2, 2), padx=(0, 10))

        # Create label and entry field for item[39]
        ttk.Label(frame1, text=self.parameters[40]).grid(row=1, column=5, pady=(2, 2), padx=(5, 5), )
        ttk.Entry(frame1, width=15).grid(row=1, column=6, pady=(2, 2), padx=(0, 10))

        # Compartments Box
        for i, entry in enumerate(self.parameters[6:9]):
            label_text = ['Comp VOL', 'NH', 'OH', ][i]
            label = ttk.Label(frame2, text=label_text)
            label.grid(row=i, column=0, columnspan=2, pady=(2, 2), padx=(10, 5), sticky=tk.E)

            entry_field = ttk.Entry(frame2, width=8)
            entry_field.grid(row=i, column=1, columnspan=3, pady=(2, 2), padx=(5, 10))
            self.value_entry[entry] = entry_field

        for i, entry in enumerate(self.parameters[9:11]):
            label_text = ['Final UH', 'First UH'][i]
            label = ttk.Label(frame2, text=label_text)
            label.grid(row=i + 3, column=0, pady=(2, 2), padx=(10, 5), sticky=tk.E)

            entry_field = ttk.Entry(frame2, width=8)
            entry_field.grid(row=i + 3, column=1, pady=(2, 2), padx=(5, 10))
            self.value_entry[entry] = entry_field

        for i, entry in enumerate(self.parameters[11:13]):
            label_text = ['Final LH', 'First LH'][i]
            label = ttk.Label(frame2, text=label_text)
            label.grid(row=i + 3, column=2, pady=(2, 2), padx=(10, 5), sticky=tk.E)

            entry_field = ttk.Entry(frame2, width=8)
            entry_field.grid(row=i + 3, column=3, pady=(2, 2), padx=(5, 10))
            self.value_entry[entry] = entry_field

            # Comp2
        for i, entry in enumerate(self.parameters[13:16]):
            label_text = ['Comp VOL', 'NH', 'OH'][i]
            label = ttk.Label(frame3, text=label_text)
            label.grid(row=i, column=0, columnspan=2, pady=(2, 2), padx=(10, 5), sticky=tk.E)

            entry_field = ttk.Entry(frame3, width=8)
            entry_field.grid(row=i, column=1, columnspan=3, pady=(2, 2), padx=(5, 10))
            self.value_entry[entry] = entry_field

        for i, entry in enumerate(self.parameters[16:18]):
            label_text = ['Final UH', 'First UH'][i]
            label = ttk.Label(frame3, text=label_text)
            label.grid(row=i + 3, column=0, pady=(2, 2), padx=(10, 5), sticky=tk.E)

            entry_field = ttk.Entry(frame3, width=8)
            entry_field.grid(row=i + 3, column=1, pady=(2, 2), padx=(5, 10))
            self.value_entry[entry] = entry_field

        for i, entry in enumerate(self.parameters[18:20]):
            label_text = ['Final LH', 'First LH'][i]
            label = ttk.Label(frame3, text=label_text)
            label.grid(row=i + 3, column=2, pady=(2, 2), padx=(10, 5), sticky=tk.E)

            entry_field = ttk.Entry(frame3, width=8)
            entry_field.grid(row=i + 3, column=3, pady=(2, 2), padx=(5, 10))
            self.value_entry[entry] = entry_field

        # Comp3
        for i, entry in enumerate(self.parameters[20:23]):
            label_text = ['Comp VOL', 'NH', 'OH'][i]
            label = ttk.Label(frame4, text=label_text)
            label.grid(row=i, column=0, columnspan=2, pady=(2, 2), padx=(10, 5), sticky=tk.E)

            entry_field = ttk.Entry(frame4, width=8)
            entry_field.grid(row=i, column=1, columnspan=3, pady=(2, 2), padx=(5, 10))
            self.value_entry[entry] = entry_field

        for i, entry in enumerate(self.parameters[23:25]):
            label_text = ['Final UH', 'First UH'][i]
            label = ttk.Label(frame4, text=label_text)
            label.grid(row=i + 3, column=0, pady=(2, 2), padx=(10, 5), sticky=tk.E)

            entry_field = ttk.Entry(frame4, width=8)
            entry_field.grid(row=i + 3, column=1, pady=(2, 2), padx=(5, 10))
            self.value_entry[entry] = entry_field

        for i, entry in enumerate(self.parameters[25:27]):
            label_text = ['Final LH', 'First LH'][i]
            label = ttk.Label(frame4, text=label_text)
            label.grid(row=i + 3, column=2, pady=(2, 2), padx=(10, 5), sticky=tk.E)

            entry_field = ttk.Entry(frame4, width=8)
            entry_field.grid(row=i + 3, column=3, pady=(2, 2), padx=(5, 10))
            self.value_entry[entry] = entry_field

        # comp4
        for i, entry in enumerate(self.parameters[27:30]):
            label_text = ['Comp VOL', 'NH', 'OH'][i]
            label = ttk.Label(frame5, text=label_text)
            label.grid(row=i, column=0, columnspan=2, pady=(2, 2), padx=(10, 5), sticky=tk.E)

            entry_field = ttk.Entry(frame5, width=8)
            entry_field.grid(row=i, column=1, columnspan=3, pady=(2, 2), padx=(5, 10))
            self.value_entry[entry] = entry_field

        for i, entry in enumerate(self.parameters[30:32]):
            label_text = ['Final UH', 'First UH'][i]
            label = ttk.Label(frame5, text=label_text)
            label.grid(row=i + 3, column=0, pady=(2, 2), padx=(10, 5), sticky=tk.E)

            entry_field = ttk.Entry(frame5, width=8)
            entry_field.grid(row=i + 3, column=1, pady=(2, 2), padx=(5, 10))
            self.value_entry[entry] = entry_field

        for i, entry in enumerate(self.parameters[32:34]):
            label_text = ['Final LH', 'First LH'][i]
            label = ttk.Label(frame5, text=label_text)
            label.grid(row=i + 3, column=2, pady=(2, 2), padx=(10, 5), sticky=tk.E)

            entry_field = ttk.Entry(frame5, width=8)
            entry_field.grid(row=i + 3, column=3, pady=(2, 2), padx=(5, 10))
            self.value_entry[entry] = entry_field

        # Truck Parameter
        ttk.Label(frame6, text='Turn Table Height').grid(row=0, column=1, pady=(5, 2))
        for i, entry in enumerate(self.parameters[34:36]):
            label_text = ['Before', 'After'][i]
            label = ttk.Label(frame6, text=label_text)
            label.grid(row=i + 1, column=0, pady=(5, 2), padx=(10, 5), sticky=tk.E)

            entry_field = ttk.Entry(frame6, width=8)
            entry_field.grid(row=i + 1, column=1, pady=(5, 2), padx=(5, 10))
            self.value_entry[entry] = entry_field

        ttk.Label(frame6, text='Balloon Height').grid(row=0, column=3, pady=(5, 2))
        for i, entry in enumerate(self.parameters[36:38]):
            label_text = ['Before', 'After'][i]
            label = ttk.Label(frame6, text=label_text)
            label.grid(row=i + 1, column=2, pady=(5, 2), padx=(10, 5), sticky=tk.E)

            entry_field = ttk.Entry(frame6, width=8)
            entry_field.grid(row=i + 1, column=3, pady=(5, 2), padx=(5, 10))
            self.value_entry[entry] = entry_field

        for i, entry in enumerate(self.parameters[38:40]):
            label_text = ['B-Spring No', 'F-Spring No'][i]
            label = ttk.Label(frame6, text=label_text)
            label.grid(row=i + 1, column=4, pady=(5, 2), padx=(10, 5), sticky=tk.E)

            entry_field = ttk.Entry(frame6, width=8)
            entry_field.grid(row=i + 1, column=5, pady=(5, 2), padx=(5, 10))
            self.value_entry[entry] = entry_field

        # Buttons
        ttk.Button(button_frame, text="Preview Entries", command=self.preview).grid(row=0, column=0, padx=5)

        save_to_excel_btn = ttk.Button(button_frame, text="Save", command=self.save_to_excel)
        save_to_excel_btn.grid(row=0, column=2, padx=5)

        reset_form_btn = ttk.Button(button_frame, text="Refresh All", command=self.reset_fields)
        reset_form_btn.grid(row=0, column=3, padx=5)

        save_to_db_btn = ttk.Button(button_frame, text="Upload to Database", command=self.show_upload_warning)
        save_to_db_btn.grid(row=0, column=1, padx=5)

    def show_upload_warning(self):
        message = "Verify your Entries before you Upload.\nYou won't be able to edit the values once saved.\nDo You Want to Save Now?"
        result = messagebox.askyesno("Warning", message)

        if result:
            self.save_to_db()
        else:
            self.preview()

    # def show_warning_dialog(self):
    #     dialog = tk.Toplevel(self.Calpage)
    #     dialog.title("Warning")
    #
    #     message = "Verify your Entries before You Upload.\nYou won't be able to edit the values once saved."
    #     label = ttk.Label(dialog, text=message)
    #     label.pack(padx=10, pady=10)
    #
    #     button_frame = ttk.Frame(dialog)
    #     button_frame.pack(padx=10, pady=10)
    #
    #     preview_button = ttk.Button(button_frame, text="Preview Now", command=self.preview)
    #     preview_button.grid(row=0, column=0, padx=5)
    #
    #     upload_button = ttk.Button(button_frame, text="Upload Now", command=self.save_to_db)
    #     upload_button.grid(row=0, column=1, padx=5)

    def save_to_db(self):
        # Connect to the database
        conn = sqlite3.connect('data.db')
        c = conn.cursor()

        try:
            # Create a table with item names as columns
            c.execute('''CREATE TABLE IF NOT EXISTS calibration_data (Date TEXT, Truck_Number TEXT, Calibrator TEXT, Transporter TEXT, 
            Total_Volume INTEGER, Chassis_Number TEXT, Comp1_Vol INTEGER, Comp1_NH INTEGER, 
            Comp1_OH INTEGER, Comp1_Final_UH INTEGER, Comp1_First_UH INTEGER, Comp1_Final_LH INTEGER, Comp1_First_LH 
            INTEGER, Comp2_Vol INTEGER, Comp2_NH INTEGER, Comp2_OH INTEGER, Comp2_Final_UH INTEGER, Comp2_First_UH 
            INTEGER, Comp2_Final_LH INTEGER, Comp2_First_LH INTEGER, Comp3_Vol INTEGER, Comp3_NH INTEGER, 
            Comp3_OH INTEGER, Comp3_Final_UH INTEGER, Comp3_First_UH INTEGER, Comp3_Final_LH INTEGER, Comp3_First_LH 
            INTEGER, Comp4_Vol INTEGER, Comp4_NH INTEGER, Comp4_OH INTEGER, Comp4_Final_UH INTEGER, Comp4_First_UH 
            INTEGER, Comp4_Final_LH INTEGER, Comp4_First_LH INTEGER, T_Table_Initial TEXT, T_Table_Final TEXT, 
            Balloon_Before TEXT, Balloon_After TEXT, B_Spring_No TEXT, Front_Spring_Number TEXT, Calibration_Status TEXT)''')
            # Insert the entry values into the table
            values = []
            for entry, entry_value in self.value_entry.items():
                values.append(entry_value.get())

            c.execute(
                "INSERT INTO calibration_data VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,"
                "?,?,?,?,?)",
                tuple(values))

            # Commit the changes and close the connection
            conn.commit()
            conn.close()

            messagebox.showinfo("Success", "Data saved to database.")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def save_to_excel(self):
        # Get the document folder path
        document_folder = os.path.expanduser("~/Documents")

        # Construct the file path for the Excel file
        file_path = os.path.join(document_folder, "Calibration Record.xlsx")

        try:
            # Load the existing Excel file if it exists, otherwise create a new DataFrame
            if os.path.isfile(file_path):
                df = pd.read_excel(file_path)
                is_existing_file = True
            else:
                df = pd.DataFrame()
                is_existing_file = False

            # Create a dictionary to store the captured values
            captured_values = {}

            # Iterate over the items [:49] and store their corresponding values
            for entry in self.parameters[:40]:
                value = self.value_entry[entry].get()
                captured_values[entry] = value

            # Create a new DataFrame from the captured values
            new_row = pd.DataFrame([captured_values])

            # Concatenate the existing DataFrame with the new row DataFrame
            df = pd.concat([df, new_row], ignore_index=True)

            # Save the DataFrame to the Excel file
            df.to_excel(file_path, index=False)

            # Display a message based on whether it's a new file or an existing file update
            if is_existing_file:
                messagebox.showinfo("Success", f"File '{file_path}' updated successfully.")
            else:
                messagebox.showinfo("Success", f"New file '{file_path}' created successfully.")
        except Exception as e:
            # Display an error message
            messagebox.showerror("Error", f"Error occurred while saving the file: {e}")

    def reset_fields(self):
        for entry in self.value_entry.values():
            entry.delete(0, 'end')

    def preview(self):
        # Disable interaction with other windows
        self.Calpage.grab_set()

        preview_window = tk.Toplevel(self.Calpage)
        preview_window.title("Preview Page")
        preview_window.minsize(width=500, height=400)
        ttk.Label(preview_window, text="Preview Your Entry Data:").pack()

        # Center the preview_window in the master window
        preview_window.update_idletasks()
        w = preview_window.winfo_width()
        h = preview_window.winfo_height()
        x = (self.Calpage.winfo_width() - w) // 2 + self.Calpage.winfo_x()
        y = (self.Calpage.winfo_height() - h) // 2 + self.Calpage.winfo_y()
        preview_window.geometry('{}x{}+{}+{}'.format(w, h, x, y))

        mainframe = tk.Frame(preview_window)
        mainframe.pack()

        preview_frame = tk.LabelFrame(mainframe, text='Input Summary')
        preview_frame.grid(row=0, column=0)

        row_count = 1
        for i, entry in enumerate(self.parameters):
            if self.value_entry[entry].get():
                ttk.Label(preview_frame, text=entry).grid(row=row_count, column=0, padx=20, sticky=tk.E)
                ttk.Label(preview_frame, text=self.value_entry[entry].get()).grid(row=row_count, column=1, padx=20,
                                                                                  sticky=tk.W)
                row_count += 1

        def close_window():
            # Release the grab on the master window
            self.Calpage.grab_release()
            preview_window.destroy()

        preview_window.protocol("WM_DELETE_WINDOW",
                                lambda: messagebox.showerror("Error", "Please close the preview window first!"))
        close_button = tk.Button(preview_window, text="CLOSE", command=close_window)
        close_button.pack()

        # Make the preview window modal
        preview_window.transient(self.Calpage)
        preview_window.wait_visibility()
        preview_window.grab_set()


#main page

    # def load_excel_file(self):
    #     documents_folder = os.path.expanduser("~/Documents")
    #     file_path = os.path.join(documents_folder, "Calibration Record.xlsx")
    #     if os.path.isfile(file_path):
    #         workbook = openpyxl.load_workbook(file_path)
    #         sheet = workbook.active
    #         data = []
    #         for row in sheet.iter_rows(values_only=True):
    #             data.append("\t".join(str(cell) for cell in row))
    #         self.display_file_contents("\n".join(data))
    #     else:
    #         self.display_file_contents("File not found.")

# def display_file_contents(self, content):
#         self.view_area.delete("1.0", tk.END)
#         self.view_area.insert(tk.END, content)



caseframe = ttk.LabelFrame(master)
caseframe.grid(row=0, column=0, sticky="ew", columnspan=2)
foot = ttk.LabelFrame(master)
foot.grid(row=3, column=0, pady=10)
inner1 = ttk.LabelFrame(caseframe)
inner1.grid(row=2, column=0, padx=(20, 80))
inner2 = ttk.LabelFrame(caseframe)
inner2.grid(row=2, column=4, padx=(80, 20))
#tEXT
ttk.Label(caseframe, text='Calibration Tracker App').grid(row=0, column=3,)
ttk.Label(foot, text='Calibration Tracker App').grid(row=0, column=0,)
ttk.Label(inner2, text='Search Truck Calibration Record.').grid(row=0, column=0, columnspan=3)
ttk.Label(inner2, text='Enter the Truck Number in Capital Letters. (e.g GH01AB)').grid(row=2, column=0, columnspan=3)

#Buttons
ttk.Button(inner1, text="Enter New Record", command=open_top_level).grid(row=1, column=0, padx=10, pady=10, sticky='w')
# Create the load button
load_button = ttk.Button(inner1, text="Load Excel File", command=load_excel_file)
load_button.grid(row=1, column=1, pady=10, padx=5)
# Create the view record button
view_button = ttk.Button(inner1, text="View Record", command=view_record)
view_button.grid(row=1, column=2, padx=5)
# Create the search button
search_button = ttk.Button(inner2, text="Search", command=search_keyword)
search_button.grid(row=1, column=2, padx=5)
# Create the search bar
search_entry = ttk.Entry(inner2, width=10)
search_entry.grid(row=1, column=1, padx=(5, 0))
# Create the search option combo box
search_option = ttk.Combobox(inner2, values=["Select Search Location", "Search from Local", "Search from Database"])
search_option.current(0)  # Set the default option
search_option.grid(row=1, column=0, padx=(5, 0))

# Increase the height of the treeview area
treeview_height = height - caseframe.winfo_height() - 2 * 10

# Create the treeview
treeview = ttk.Treeview(master)
treeview.grid(row=1, column=0, sticky="nsew", padx=10, pady=10, rowspan=2)



# Create the vertical scrollbar
vsb = ttk.Scrollbar(master, orient="vertical", command=treeview.yview)
vsb.grid(row=1, column=1, sticky="ns")
treeview.configure(yscrollcommand=vsb.set)

# Create the horizontal scrollbar
hsb = ttk.Scrollbar(master, orient="horizontal", command=treeview.xview)
hsb.grid(row=2, column=0, sticky="ew")
treeview.configure(xscrollcommand=hsb.set)

# Configure grid weights
master.grid_rowconfigure(0, weight=0)  # Adjust the weight as needed
master.grid_rowconfigure(1, weight=1)  # Adjust the weight as needed
master.grid_columnconfigure(0, weight=1)

HEADINGS = ['Date', 'Truck Number', 'Calibrator', 'Transporter', 'Total Volume', 'Chassis Number',
            'Comp 1 Vol', 'NH', 'OH', 'Final UH', 'First UH', 'Final LH', 'First LH',
            'Comp2 Vol', 'NH', 'OH', 'Final UH', 'First UH', 'Final LH', 'First LH',
            'Comp3 Vol', 'NH', 'OH', 'Final UH', 'First UH', 'Final LH', 'First LH',
            'Comp4 Vol', 'NH', 'OH', 'Final UH', 'First UH', 'Final LH', 'First LH']

master.mainloop()
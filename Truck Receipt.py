import tkinter as tk
# import ttkbootstrap
from tkinter import ttk
from tkcalendar import DateEntry
from tkinter import filedialog, messagebox
import pandas as pd
import os
import sqlite3


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

def shortage_page():
    top_level = tk.Toplevel(master)
    ScoreEntryForm(top_level)


def load_excel_file():
    file_path = os.path.expanduser("~/Documents/Truck Receipt Record.xlsx")

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

    # Display the first five headings and values horizontally
    for i, heading in enumerate(treeview['columns'][:5]):
        ttk.Label(top_frame, text=heading).grid(row=0, column=i, padx=5, pady=5)
        ttk.Label(top_frame, text=values[i]).grid(row=1, column=i, padx=5, pady=5)

    # Create a frame for the bottom section (vertical display)
    bottom_frame = ttk.Frame(popup)
    bottom_frame.pack(padx=10, pady=(20, 10))  # Add some gap between the sections

    # Display the remaining headings and values vertically, 3 items per column
    for i, heading in enumerate(treeview['columns'][5:]):
        row = i % 3
        column = i // 3

        ttk.Label(bottom_frame, text=heading).grid(row=row, column=column * 2, sticky='e', padx=5, pady=5)
        ttk.Label(bottom_frame, text=values[i+5]).grid(row=row, column=column * 2 + 1, sticky='w', padx=5, pady=5)

    popup.mainloop()

def search_keyword():
    keyword = search_entry.get()

    if not keyword:
        messagebox.showwarning("Warning", "Please enter a keyword for search.")
        return

    file_path = os.path.expanduser("~/Documents/Calibration Record.xlsx")

    try:
        df = pd.read_excel(file_path)
    except:
        messagebox.showerror("Error", "Unable to access the file!")
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
    def __init__(self, master):
        self.master = master
        self.master.title("Product Receipt DB")
        width = 900
        height = 600
        screen_width = self.master.winfo_screenwidth()
        screen_height = self.master.winfo_screenheight()
        x = int((screen_width / 2) - (width / 2))
        y = int((screen_height / 2) - (height / 2))
        self.master.geometry(f"{width}x{height}+{x}+{y}")

        self.parameters = ['Date', 'Truck Number', 'Transporter', 'Truck Total Volume', 'Received Location',
                           'Comp1 Loaded-Vol', 'Comp1 Loaded-UH', 'Comp1 Loaded-LH', 'Comp1 Arrival-Vol', 'Comp1 Arrival-UH)', 'Comp1 Arrival-lH',
                           'Comp2 Loaded-Vol', 'Comp2 Loaded-UH', 'Comp2 Loaded-LH', 'Comp2 Arrival-Vol', 'Comp2 Arrival-UH)', 'Comp2 Arrival-lH',
                           'Comp3 Loaded-Vol', 'Comp3 Loaded-UH', 'Comp3 Loaded-LH', 'Comp3 Arrival-Vol', 'Comp3 Arrival-UH)', 'Comp3 Arrival-lH',
                           'Comp4 Loaded-Vol', 'Comp4 Loaded-UH', 'Comp4 Loaded-LH', 'Comp4 Arrival-Vol', 'Comp4 Arrival-UH)', 'Comp4 Arrival-lH',
                           ]

        self.value_entry = {}
        for entry in self.parameters:
            self.value_entry[entry] = ttk.Entry(self.master)
        # FRAMES
        caseframe = ttk.Frame(self.master)
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
            elif entry == 'Transporter':
                self.value_entry[entry] = ttk.Combobox(frame1, values=['Ola', 'Olu', 'Ade', 'Ife', 'Ire'], width=10)
                self.value_entry[entry].grid(row=0, column=i * 2 + 1, pady=(10, 2), padx=(0, 10))
            else:
                self.value_entry[entry] = ttk.Entry(frame1, width=10)
                self.value_entry[entry].grid(row=0, column=i * 2 + 1, pady=(10, 2), padx=(0, 10))




        # Compartments Box

        ttk.Label(frame2, text='Loading Details').grid(row=0, column=1, pady=(5, 2))
        for i, entry in enumerate(self.parameters[5:8]):
            label_text = ['Loaded-Vol','UH', 'LH'][i]
            label = ttk.Label(frame2, text=label_text)
            label.grid(row=i + 1, column=0, pady=(5, 2), padx=(10, 5), sticky=tk.E)

            entry_field = ttk.Entry(frame2, width=8)
            entry_field.grid(row=i + 1, column=1, pady=(5, 2), padx=(5, 10))
            self.value_entry[entry] = entry_field

        ttk.Label(frame2, text='Arrival Details').grid(row=0, column=3, pady=(5, 2))
        for i, entry in enumerate(self.parameters[8:11]):
            label_text = ['Arrival-Vol','UH', 'LH'][i]
            label = ttk.Label(frame2, text=label_text)
            label.grid(row=i + 1, column=2, pady=(5, 2), padx=(10, 5), sticky=tk.E)

            entry_field = ttk.Entry(frame2, width=8)
            entry_field.grid(row=i + 1, column=3, pady=(5, 2), padx=(5, 10))
            self.value_entry[entry] = entry_field

            # Comp2
        ttk.Label(frame3, text='Loading Details').grid(row=0, column=1, pady=(5, 2))
        for i, entry in enumerate(self.parameters[11:14]):
            label_text = ['Loaded-Vol','UH', 'LH'][i]
            label = ttk.Label(frame3, text=label_text)
            label.grid(row=i + 1, column=0, pady=(5, 2), padx=(10, 5), sticky=tk.E)

            entry_field = ttk.Entry(frame3, width=8)
            entry_field.grid(row=i + 1, column=1, pady=(5, 2), padx=(5, 10))
            self.value_entry[entry] = entry_field

        ttk.Label(frame3, text='Arrival Details').grid(row=0, column=3, pady=(5, 2))
        for i, entry in enumerate(self.parameters[14:17]):
            label_text = ['Arrival-Vol','UH', 'LH'][i]
            label = ttk.Label(frame3, text=label_text)
            label.grid(row=i + 1, column=2, pady=(5, 2), padx=(10, 5), sticky=tk.E)

            entry_field = ttk.Entry(frame3, width=8)
            entry_field.grid(row=i + 1, column=3, pady=(5, 2), padx=(5, 10))
            self.value_entry[entry] = entry_field


        # Comp3
        ttk.Label(frame4, text='Loading Details').grid(row=0, column=1, pady=(5, 2))
        for i, entry in enumerate(self.parameters[17:20]):
            label_text = ['Loaded-Vol','UH', 'LH'][i]
            label = ttk.Label(frame4, text=label_text)
            label.grid(row=i + 1, column=0, pady=(5, 2), padx=(10, 5), sticky=tk.E)

            entry_field = ttk.Entry(frame4, width=8)
            entry_field.grid(row=i + 1, column=1, pady=(5, 2), padx=(5, 10))
            self.value_entry[entry] = entry_field

        ttk.Label(frame4, text='Arrival Details').grid(row=0, column=3, pady=(5, 2))
        for i, entry in enumerate(self.parameters[20:23]):
            label_text = ['Arrival-Vol', 'UH', 'LH'][i]
            label = ttk.Label(frame4, text=label_text)
            label.grid(row=i + 1, column=2, pady=(5, 2), padx=(10, 5), sticky=tk.E)

            entry_field = ttk.Entry(frame4, width=8)
            entry_field.grid(row=i + 1, column=3, pady=(5, 2), padx=(5, 10))
            self.value_entry[entry] = entry_field

        # comp4
        ttk.Label(frame5, text='Loading Details').grid(row=0, column=1, pady=(5, 2))
        for i, entry in enumerate(self.parameters[23:26]):
            label_text = ['Loaded-Vol','UH', 'LH'][i]
            label = ttk.Label(frame5, text=label_text)
            label.grid(row=i + 1, column=0, pady=(5, 2), padx=(10, 5), sticky=tk.E)

            entry_field = ttk.Entry(frame5, width=8)
            entry_field.grid(row=i + 1, column=1, pady=(5, 2), padx=(5, 10))
            self.value_entry[entry] = entry_field

        ttk.Label(frame5, text='Arrival Details').grid(row=0, column=3, pady=(5, 2))
        for i, entry in enumerate(self.parameters[26:29]):
            label_text = ['Arrival-Vol','UH', 'LH'][i]
            label = ttk.Label(frame5, text=label_text)
            label.grid(row=i + 1, column=2, pady=(5, 2), padx=(10, 5), sticky=tk.E)

            entry_field = ttk.Entry(frame5, width=8)
            entry_field.grid(row=i + 1, column=3, pady=(5, 2), padx=(5, 10))
            self.value_entry[entry] = entry_field

        # Truck Parameter
        # ttk.Label(frame6, text='Turn Table Height').grid(row=0, column=1, pady=(5, 2))
        # for i, entry in enumerate(self.parameters[34:36]):
        #     label_text = ['Before', 'After'][i]
        #     label = ttk.Label(frame6, text=label_text)
        #     label.grid(row=i + 1, column=0, pady=(5, 2), padx=(10, 5), sticky=tk.E)
        #
        #     entry_field = ttk.Entry(frame6, width=8)
        #     entry_field.grid(row=i + 1, column=1, pady=(5, 2), padx=(5, 10))
        #     self.value_entry[entry] = entry_field
        #
        # ttk.Label(frame6, text='Balloon Height').grid(row=0, column=3, pady=(5, 2))
        # for i, entry in enumerate(self.parameters[36:38]):
        #     label_text = ['Before', 'After'][i]
        #     label = ttk.Label(frame6, text=label_text)
        #     label.grid(row=i + 1, column=2, pady=(5, 2), padx=(10, 5), sticky=tk.E)
        #
        #     entry_field = ttk.Entry(frame6, width=8)
        #     entry_field.grid(row=i + 1, column=3, pady=(5, 2), padx=(5, 10))
        #     self.value_entry[entry] = entry_field
        #
        # for i, entry in enumerate(self.parameters[38:40]):
        #     label_text = ['B-Spring No', 'F-Spring No'][i]
        #     label = ttk.Label(frame6, text=label_text)
        #     label.grid(row=i + 1, column=4, pady=(5, 2), padx=(10, 5), sticky=tk.E)
        #
        #     entry_field = ttk.Entry(frame6, width=8)
        #     entry_field.grid(row=i + 1, column=5, pady=(5, 2), padx=(5, 10))
        #     self.value_entry[entry] = entry_field

        # Buttons
        ttk.Button(button_frame, text="Preview Entries", command=self.preview).grid(row=0, column=0, padx=5)

        save_to_excel_btn = ttk.Button(button_frame, text="Save to Excel", command=self.save_to_excel)
        save_to_excel_btn.grid(row=0, column=2, padx=5)

        reset_form_btn = ttk.Button(button_frame, text="Refresh All", command=self.reset_fields)
        reset_form_btn.grid(row=0, column=3, padx=5)

        save_to_db_btn = ttk.Button(button_frame, text="Upload to Database", command=self.save_to_db)
        save_to_db_btn.grid(row=0, column=1, padx=5)



    def save_to_db(self):
        # Connect to the database
        conn = sqlite3.connect('Receipt-data.db')
        c = conn.cursor()

        try:
            # Create a table with item names as columns
            c.execute('''CREATE TABLE IF NOT EXISTS receipt-data (
                  date TEXT,
                  truck_number TEXT,
                  transporter TEXT,
                  truck_total_volume REAL,
                  received_location TEXT,
                  comp1_loaded_vol REAL,
                  comp1_loaded_uh INTEGER,
                  comp1_loaded_lh INTEGER,
                  comp1_arrival_vol REAL,
                  comp1_arrival_uh INTEGER,
                  comp1_arrival_lh INTEGER,
                  comp2_loaded_vol REAL,
                  comp2_loaded_uh INTEGER,
                  comp2_loaded_lh INTEGER,
                  comp2_arrival_vol REAL,
                  comp2_arrival_uh INTEGER,
                  comp2_arrival_lh INTEGER,
                  comp3_loaded_vol REAL,
                  comp3_loaded_uh INTEGER,
                  comp3_loaded_lh INTEGER,
                  comp3_arrival_vol REAL,
                  comp3_arrival_uh INTEGER,
                  comp3_arrival_lh INTEGER,
                  comp4_loaded_vol REAL,
                  comp4_loaded_uh INTEGER,
                  comp4_loaded_lh INTEGER,
                  comp4_arrival_vol REAL,
                  comp4_arrival_uh INTEGER,
                  comp4_arrival_lh INTEGER)''')
            # Insert the entry values into the table
            values = []
            for entry, entry_value in self.value_entry.items():
                values.append(entry_value.get())

            c.execute(
                "INSERT INTO calibration_data VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, "
                "?, ?, ?, ?, ?, ?)",
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
        file_path = os.path.join(document_folder, "Truck Receipt Record.xlsx")

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
            for entry in self.parameters[:30]:
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
        self.master.grab_set()

        preview_window = tk.Toplevel(self.master)
        preview_window.title("Preview Page")
        preview_window.minsize(width=500, height=400)
        ttk.Label(preview_window, text="Preview Your Entry Data:").pack()

        # Center the preview_window in the master window
        preview_window.update_idletasks()
        w = preview_window.winfo_width()
        h = preview_window.winfo_height()
        x = (self.master.winfo_width() - w) // 2 + self.master.winfo_x()
        y = (self.master.winfo_height() - h) // 2 + self.master.winfo_y()
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
            self.master.grab_release()
            preview_window.destroy()

        preview_window.protocol("WM_DELETE_WINDOW",
                                lambda: messagebox.showerror("Error", "Please close the preview window first!"))
        close_button = tk.Button(preview_window, text="CLOSE", command=close_window)
        close_button.pack()

        # Make the preview window modal
        preview_window.transient(self.master)
        preview_window.wait_visibility()
        preview_window.grab_set()




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
    #     self.view_area.delete("1.0", tk.END)
    #     self.view_area.insert(tk.END, content)


class ScoreEntryForm:
    def __init__(self, master):
        self.master = master
        self.master.title("Shortage Calculator")
        width = 1000
        height = 500
        screen_width = self.master.winfo_screenwidth()
        screen_height = self.master.winfo_screenheight()
        x = int((screen_width / 2) - (width / 2))
        y = int((screen_height / 2) - (height / 2))
        self.master.geometry(f"{width}x{height}+{x}+{y}")

        self.entry_values = ['Date', 'Truck Number', 'Transporter', 'Truck Total Volume', 'Received Location', 'Product Type',
                           'Comp1 Vol', 'Comp1 Loaded-UH', 'Comp1 Loaded-LH', 'Comp1 Arrival-UH', 'Comp1 Arrival-LH', 'Variance',
                           'Comp2 Vol', 'Comp2 Loaded-UH', 'Comp2 Loaded-LH', 'Comp2 Arrival-UH', 'Comp2 Arrival-LH', 'Variance',
                           'Comp3 Vol', 'Comp3 Loaded-UH', 'Comp3 Loaded-LH', 'Comp3 Arrival-UH', 'Comp3 Arrival-LH', 'Variance',
                           'Comp4 Vol', 'Comp4 Loaded-UH', 'Comp4 Loaded-LH', 'Comp4 Arrival-UH', 'Comp4 Arrival-LH', 'Variance',
                           'Comp1 Tank_B4', 'Comp1 Tank_After', 'Comp1 Tank-Vol', 'Comp2 Tank_B4', 'Comp2 Tank_After', 'Comp2 Tank-Vol',
                           'Comp3 Tank_B4', 'Comp3 Tank_After', 'Comp3 Tank-Vol', 'Comp4 Tank_B4', 'Comp4 Tank_After', 'Comp4 Tank-Vol']

        self.value_entry = {}
        for entry in self.entry_values:
            self.value_entry[entry] = ttk.Entry(self.master)
        # FRAMES
        caseframe = ttk.Frame(self.master)
        caseframe.pack()
        frame1 = tk.LabelFrame(caseframe, text='Truck Details')
        frame1.grid(row=0, column=0, columnspan=2, sticky=tk.W)
        frame2 = tk.LabelFrame(caseframe, text='Compartments Values')
        frame2.grid(row=2, column=0, columnspan=2)
        frame3 = tk.LabelFrame(caseframe, text='ULLAGE Result Area')
        frame3.grid(row=3, column=0, padx=10, pady=10)
        frame5 = tk.LabelFrame(caseframe, text='LIQUID HEIGHT Result Area')
        frame5.grid(row=3, column=1, padx=10, pady=10)
        frame4 = tk.LabelFrame(caseframe, text='Product')
        frame4.grid(row=1, column=0, columnspan=2, pady=10)
        button_frame = ttk.Frame(caseframe)
        button_frame.grid(row=4, column=0, columnspan=2, pady=10)

        for i, entry in enumerate(self.entry_values[:5]):
            label = ttk.Label(frame1, text=entry)
            label.grid(row=0, column=i * 2, pady=(10, 2), padx=(10, 5), sticky=tk.E)

            if entry == 'Date':
                # Make the Date entry a calendar date picker
                self.value_entry[entry] = DateEntry(frame1, width=12, background='darkblue',
                                                    foreground='white', borderwidth=2, year=2023,
                                                    date_pattern='dd/MM/yyyy')
                self.value_entry[entry].grid(row=0, column=i * 2 + 1, pady=(10, 2), padx=(0, 10))
            elif entry == 'Product':
                self.value_entry[entry] = ttk.Combobox(frame1, values=['Ola', 'Olu', 'Ade', 'Ife', 'Ire'], width=10)
                self.value_entry[entry].grid(row=0, column=i * 2 + 1, pady=(10, 2), padx=(0, 10))
            else:
                self.value_entry[entry] = ttk.Entry(frame1, width=10)
                self.value_entry[entry].grid(row=0, column=i * 2 + 1, pady=(10, 2), padx=(0, 10))

        # Create the search option combo box
        search_option = ttk.Combobox(frame4, values=["Select", "Ullage Parameters", "Liquid Height Parameters"])
        search_option.current(0)  # Set the default option
        search_option.grid(row=1, column=0, padx=(5, 0))

        # Create the search button
        search_button = ttk.Button(frame4, text="Load Parameters")
        search_button.grid(row=1, column=2, padx=5)
        # Create the search bar
        search_entry = ttk.Entry(frame4, width=10)
        search_entry.grid(row=1, column=1, padx=(5, 0))

        column_headings = ['Vol', 'Loaded-UH', 'Loaded-LH', 'Arrival-UH', 'Arrival-LH', 'Variance']
        row_labels = ['Comp1', 'Comp2', 'Comp3', 'Comp4']

        # Create the entry fields and column headings
        entry_fields = []
        headings = []

        # Create the entry fields and column headings for rows 1 to 4
        for i, item in enumerate(self.entry_values[:24]):
            row = i // 6  # Rows 1 to 4
            column = i % 6  # Columns 0 to 5

            # Create the column heading
            heading_label = ttk.Label(frame2, text=column_headings[column])
            heading_label.grid(row=0, column=column + 1, padx=5, pady=5)
            headings.append(heading_label)

            # Create the entry field
            entry = ttk.Entry(frame2, width=10)
            entry.grid(row=row + 1, column=column + 1, padx=5, pady=5)
            entry_fields.append(entry)

        # Create the row labels for rows 1 to 4
        for i, label_text in enumerate(row_labels):
            label = ttk.Label(frame2, text=label_text)
            label.grid(row=i + 1, column=0, padx=5, pady=5)


        # Buttons
        ttk.Button(button_frame, text="Preview Entries",).grid(row=0, column=0, padx=5)

        save_to_excel_btn = ttk.Button(button_frame, text="Save to Excel", )
        save_to_excel_btn.grid(row=0, column=2, padx=5)

        reset_form_btn = ttk.Button(button_frame, text="Refresh All",)
        reset_form_btn.grid(row=0, column=3, padx=5)

        calculate_button = ttk.Button(button_frame, text="Calculate Output", )
        calculate_button.grid(row=0, column=1, padx=5)

        # Create Ullage result display areas
        ttk.Label(frame3, text="Total Loaded Vol").grid(row=0, column=0, padx=5, pady=5, sticky=tk.E)
        self.total_withdrawn_display = ttk.Label(frame3, text="0")
        self.total_withdrawn_display.grid(row=0, column=1, padx=5, pady=5)

        ttk.Label(frame3, text="Total Arrival Vol").grid(row=1, column=0, padx=5, pady=5, sticky=tk.E)
        self.total_correction_display = ttk.Label(frame3, text="0")
        self.total_correction_display.grid(row=1, column=1, padx=5, pady=5)

        ttk.Label(frame3, text="Total Variance (SMR - Arrival)").grid(row=2, column=0, padx=5, pady=5, sticky=tk.E)
        self.total_additives_display = ttk.Label(frame3, text="0")
        self.total_additives_display.grid(row=2, column=1, padx=5, pady=5)

        ttk.Label(frame3, text="Variance (Chart-SMR)").grid(row=2, column=0, padx=5, pady=5, sticky=tk.E)
        self.total_additives_display = ttk.Label(frame3, text="0")
        self.total_additives_display.grid(row=2, column=1, padx=5, pady=5)

        ttk.Label(frame3, text="Variance (Chart-Arrival)").grid(row=2, column=0, padx=5, pady=5, sticky=tk.E)
        self.total_additives_display = ttk.Label(frame3, text="0")
        self.total_additives_display.grid(row=2, column=1, padx=5, pady=5)

        ttk.Label(frame3, text="REMARK").grid(row=3, column=0, padx=5, pady=5, sticky=tk.E)
        self.total_bulk_reproduced_display = ttk.Label(frame3, text="0")
        self.total_bulk_reproduced_display.grid(row=3, column=1, padx=5, pady=5)


        # Create result display areas
        ttk.Label(frame5, text="Total Loaded Vol").grid(row=0, column=0, padx=5, pady=5, sticky=tk.E)
        self.total_withdrawn_display = ttk.Label(frame5, text="0")
        self.total_withdrawn_display.grid(row=0, column=1, padx=5, pady=5)

        ttk.Label(frame5, text="Total Arrival Vol").grid(row=1, column=0, padx=5, pady=5, sticky=tk.E)
        self.total_correction_display = ttk.Label(frame5, text="0")
        self.total_correction_display.grid(row=1, column=1, padx=5, pady=5)

        ttk.Label(frame5, text="Total Variance (SMR - Arrival)").grid(row=2, column=0, padx=5, pady=5, sticky=tk.E)
        self.total_additives_display = ttk.Label(frame5, text="0")
        self.total_additives_display.grid(row=2, column=1, padx=5, pady=5)

        ttk.Label(frame5, text="Variance (Chart-SMR)").grid(row=2, column=0, padx=5, pady=5, sticky=tk.E)
        self.total_additives_display = ttk.Label(frame5, text="0")
        self.total_additives_display.grid(row=2, column=1, padx=5, pady=5)

        ttk.Label(frame5, text="Variance (Chart-Arrival)").grid(row=2, column=0, padx=5, pady=5, sticky=tk.E)
        self.total_additives_display = ttk.Label(frame5, text="0")
        self.total_additives_display.grid(row=2, column=1, padx=5, pady=5)

        ttk.Label(frame5, text="REMARK").grid(row=3, column=0, padx=5, pady=5, sticky=tk.E)
        self.total_bulk_reproduced_display = ttk.Label(frame5, text="0")
        self.total_bulk_reproduced_display.grid(row=3, column=1, padx=5, pady=5)




caseframe = ttk.LabelFrame(master)
caseframe.grid(row=0, column=0, sticky="ew", columnspan=2)
foot = ttk.LabelFrame(master)
foot.grid(row=3, column=0, pady=10)
inner1 = ttk.LabelFrame(caseframe)
inner1.grid(row=2, column=0, padx=(20, 40))
inner2 = ttk.LabelFrame(caseframe)
inner2.grid(row=2, column=4, padx=(0, 20))
#tEXT
ttk.Label(caseframe, text='Truck Discharge Tracker App').grid(row=0, column=3,)
ttk.Label(foot, text='Truck Discharge Tracker App').grid(row=0, column=0,)
ttk.Label(inner2, text='Search Truck Discharge History.').grid(row=0, column=0, columnspan=3)
ttk.Label(inner2, text='Enter the Truck Number in Capital Letters. (e.g GH01AB)').grid(row=2, column=0, columnspan=3)

#Buttons
ttk.Button(inner1, text="Enter New Record", command=open_top_level).grid(row=1, column=0, padx=10, pady=10, sticky='w')
# Create the load button
load_button = ttk.Button(inner1, text="Load Database", command=load_excel_file)
load_button.grid(row=1, column=1, pady=10, padx=5)
# Create the view record button
view_button = ttk.Button(inner1, text="View Record", command=view_record)
view_button.grid(row=1, column=2, padx=5)
# Create the shortage calculation button
shortage_page_button = ttk.Button(inner1, text="Calculate Shortage", command=shortage_page)
shortage_page_button.grid(row=1, column=3, padx=5)
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


master.mainloop()


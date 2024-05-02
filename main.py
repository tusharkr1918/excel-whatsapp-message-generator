"""
App: Excel Whatsapp Message Generator
Author: Tushar Kumar
Version: 1.0
Date: March 9, 2024
Description: A sample application for generating whatsapp messages.
"""

import os
import re
import json
import sys
import threading
import numpy as np
import pandas as pd
import customtkinter as ctk
from PIL import Image
from datetime import datetime, timedelta
from tkinter.filedialog import askdirectory, askopenfilename
from utils.generate_hyperlink import excel_column_letter_to_index, process_branch_data

ctk.set_appearance_mode("dark")

class ExcelHyperlinkSplitter:
    pattern = r"([A-Za-z]+)|(\"[^\"]*\")|(\[[A-Za-z]+\.[A-Za-z]+\])"

    DATA_DIR = os.path.join(os.path.expanduser("~"), "ExcelHyperlinkSplitter")
    os.makedirs(DATA_DIR, exist_ok=True)

    JSON_FILE = os.path.join(DATA_DIR, "state_data.json")

    @staticmethod
    def resource_path(relative_path):
        base_path = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
        return os.path.join(base_path, relative_path)

    @staticmethod
    def validate_alphabetic_input(char):
        return char.isalpha()

    @staticmethod
    def validate_numeric_input(char):
        return char.isdigit() or len(char) == 0

    def __init__(self, root):
        self.root = root
        print(title:=f"Whatsapp Message Generator for Excel - v1.0")
        self.root.title(title)
        self.root.iconbitmap(ExcelHyperlinkSplitter.resource_path(r"assets\excel-icon.ico"))

        # Set the dimensions of the window (width x height)
        self.window_width = 600
        self.window_height = 700

        # Get the screen width and height
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()

        # Calculate the x and y coordinates to center the window
        x_coordinate = (screen_width - self.window_width) // 2
        y_coordinate = (screen_height - self.window_height) // 2

        # Set the window geometry to center the window
        self.root.geometry(f"{self.window_width}x{self.window_height}+{x_coordinate}+{y_coordinate}")
        self.root.resizable(width=False, height=False)

        self.validation_1 = self.root.register(ExcelHyperlinkSplitter.validate_alphabetic_input)
        self.validation_2 = self.root.register(ExcelHyperlinkSplitter.validate_numeric_input)


        # Create the first set of Label, Entry, and Button
        self.excel_widgets = self.create_widgets(row=0, label="Select excel path", type="askopenfilename")

        # Create the second set of Label, Entry, and Button
        self.output_widgets = self.create_widgets(row=3, label="Select output directory path", type="askdirectory")

        # Adjust column weights to make the Entry boxes expandable
        self.root.columnconfigure(0, weight=1)
        self.root.columnconfigure(1, weight=1)

        self.format_frame = self.create_frame(row=5, height=100, label="Format message",state="disabled", margine_y=15)
        self.format_frame[1].bind('<<Modified>>', self.on_text_change)

        self.preview_frame = self.create_frame(row=7, height=160, label="Preview message", state="disabled", margine_y=15)
        # self.preview_entry = ctk.CTkEntry(self.root, height=28, width=86)

        # self.c = ctk.CTkCanvas(self.root, width=78, height=28+15, bg="#1D1E1E", highlightthickness=0)
        # self.c.grid(row=7, column=3, padx=(0, 20), sticky="es")

    
        self.entry_start = self.create_data_widgets(row=9)

        self.progress_bar = ctk.CTkProgressBar(self.root, mode="determinate", height=3, progress_color="grey11")
        self.progress_bar.set(1)
        self.progress_bar.grid(row=10, column=0, columnspan=4, sticky="we", pady=(0, 15), padx=15)

        self.bottom_frame = ctk.CTkFrame(self.root, height=40, fg_color="transparent")
        self.bottom_frame.grid(row=11, column=0, columnspan=2, padx=(15,0), sticky="wens")

        self.statusVar = ctk.StringVar()
        self.statusVar.set(f"[info]:")
        self.status = ctk.CTkEntry(self.bottom_frame, height=40, fg_color="white", border_width=2, border_color="orange", text_color="black", corner_radius=6, textvariable=self.statusVar, font=("Consolas", 15, "bold"), state="disabled")
        self.status.grid(row=0, column=0, sticky="ew")

        # Set weight for the status column
        self.bottom_frame.columnconfigure(0, weight=1)

        self.recover = ctk.CTkButton(self.bottom_frame, image=ctk.CTkImage(Image.open((ExcelHyperlinkSplitter.resource_path(r"assets\load.png"))), size=(20, 20)), text="", width=40, height=30, bg_color="transparent", fg_color=['#3B8ED0', '#1F6AA5'], command=self.load_state)
        self.recover.grid(row=0, column=1, padx=11, sticky="ewns")

        self.process_btn = ctk.CTkButton(self.root, text=f"START", command=self.hyperlink_splitter, height=40, width=10, font=("Helvetica", 15, "bold"))
        self.process_btn.grid(row=11, column=3, padx=(0, 15), sticky="wens")

        
    def save_state(self):
        state_data = {
            "file_path": self.file_path,
            "output_path": self.output_file_path,
            "branch_column": self.branch_column.get(),
            "phone_column": self.phone_column.get(),
            "hyperlink_column": self.hyperlink_column.get(),
            "format_string": self.format_frame[1].get("1.0", "end-1c"),
            "chunk_size": self.chunk_size.get(),
            "splitby_branch": self.branch_checkbox.get()
        }
        with open(ExcelHyperlinkSplitter.JSON_FILE, "w") as file:
            json.dump(state_data, file)

    def load_state(self):
        try:
            with open(ExcelHyperlinkSplitter.JSON_FILE, "r") as file:
                state_data = json.load(file)
                self.excel_widgets[3].set(state_data["file_path"]) 
                self.file_path = state_data["file_path"]

                self.output_widgets[3].set(state_data["output_path"])
                self.output_file_path = state_data["output_path"]

                self.load_excel_file(state_data["file_path"])

                self.format_frame[1].delete("1.0", "end")
                self.format_frame[1].insert("end", state_data["format_string"])
  
                self.phone_column.delete("0", "end")
                self.phone_column.insert("0", state_data["phone_column"])

                self.checkbox_var.set(state_data["splitby_branch"])
                self.show_chunk_entry(data=state_data["branch_column"])

                self.hyperlink_column.delete("0", "end")
                self.hyperlink_column.insert("0", state_data["hyperlink_column"])

                self.chunk_size.delete("0", "end")
                self.chunk_size.insert("0", state_data["chunk_size"])

                self.update_status("Previous state loaded successfully.", "green")
          
        except FileNotFoundError as e:
            self.update_status(str(e), "red")

    def show_chunk_entry(self, data=None):
        if not self.checkbox_var.get():
            self.label.grid(row=1, column=2, pady=(0, 15), sticky="w")
            self.chunk_size.grid(row=1, column=3, pady=(0, 15), sticky="ewns")
            
            if data:
                self.branch_column.delete("0", "end")
                self.branch_column.insert("0", data)

            self.branch_column.configure(state="disabled", fg_color="grey30")
        else:
            self.label.grid_remove()
            self.chunk_size.grid_remove()
            self.branch_column.grid(row=0, column=1, sticky="w")
            self.branch_column.configure(state="normal", fg_color=['#F9F9FA', '#343638'])

            if data:
                self.branch_column.delete("0", "end")
                self.branch_column.insert("0", data)


    def create_data_widgets(self, row):
        # Create a new frame for this row
        frame = ctk.CTkFrame(self.root, fg_color="transparent")
        frame.grid(row=row, column=0, columnspan=4, padx=15, pady=(15, 0), sticky="ew")

        # Phone Column
        label_widget = ctk.CTkLabel(frame, text="Phone column:", font=("Comic Sans MS", 16, "bold"))
        label_widget.grid(row=0, column=0, pady=(0, 15), sticky="w")

        self.phone_column = ctk.CTkEntry(frame, font=("Courier New", 14), placeholder_text="A", width=50, validate="key", validatecommand=(self.validation_1, "%S"))
        
        self.phone_column.grid(row=0, column=1, padx=(40,40), pady=(0, 15), sticky="e")

        # Hyperlink Column
        label_widget = ctk.CTkLabel(frame, text="Hyperlink column:", font=("Comic Sans MS", 16, "bold"))
        label_widget.grid(row=1, column=0, pady=(0, 15), sticky="w")

        self.hyperlink_column = ctk.CTkEntry(frame, font=("Courier New", 14), placeholder_text="AB", width=50, validate="key", validatecommand=(self.validation_1, "%S"))
        self.hyperlink_column.grid(row=1, column=1, padx=(40,40), pady=(0, 15), sticky="e")

        label_widget = ctk.CTkLabel(frame, text="Split by column value: ", font=("Comic Sans MS", 16, "bold"))
        label_widget.grid(row=0, column=2, padx=(0, 37), pady=(0, 15), sticky="w")


        self.label = ctk.CTkLabel(frame, text=f"No. of rows in each file:", font=("Comic Sans MS", 16, "bold"))

        entry_var = ctk.StringVar()
        entry_var.set(200)
        self.chunk_size = ctk.CTkEntry(frame, font=("Courier New", 14), validate="key", textvariable=entry_var, width=80, validatecommand=(self.validation_2, "%P"))

        self.branch_frame = ctk.CTkFrame(frame, fg_color="transparent", height=28)
        self.branch_frame.grid(row=0, column=3, pady=(0, 15), sticky="w")  # Align the frame to the left (west)

        self.checkbox_var = ctk.BooleanVar()
        self.checkbox_var.set(True)
        self.branch_checkbox = ctk.CTkCheckBox(self.branch_frame, text="", border_width=2, command=self.show_chunk_entry, border_color=['#979DA2', '#565B5E'], variable=self.checkbox_var, width=2, checkbox_height=28, checkbox_width=28)
        # self.branch_checkbox.
        self.branch_checkbox.grid(row=0, column=0, sticky="w")  # Align the checkbox to the left within the branch_frame

        # Create an entry box that spans the remaining available space
        self.branch_column = ctk.CTkEntry(
            self.branch_frame, border_width=2, 
            placeholder_text="B",
            width=50, validate="key",
            validatecommand=(self.validation_1, "%S"), 
            font=("Courier New", 14)
        )
        self.branch_column.grid(row=0, column=1, sticky="w")  # Align the entry box to the left


    def generate_files(self):
        branch_column_index = excel_column_letter_to_index(self.branch_column.get())
        branch_names = self.df.iloc[1:, branch_column_index].astype(str).str.strip().unique()

        for branch_name in branch_names:
            no_branch_col = process_branch_data(
                self.df,
                branch_name, 
                self.file_path,
                self.output_file_path,
                self.branch_column.get() if self.branch_checkbox.get() else "",
                self.phone_column.get(), 
                self.hyperlink_column.get(), 
                self.concat_string[:-1], 
                int(self.chunk_size.get())
            )
            if no_branch_col:
                break
            else:
                print(f"{branch_name}...done!")

        
        self.configure_progress_bar("lightgreen")
        self.update_status("Successfully generated!", "green")

        self.save_state()

    def hyperlink_splitter(self):
        if self.output_widgets[1].get() == "":
            self.update_status("Please provde the output directory", "red")
            return

        if self.phone_column.get() == "":
            self.update_status("Please provide the phone column", "red")
            return
        
        if self.hyperlink_column.get() == "":
            self.update_status("Please provide the hyperlink column", "red")
            return
        
        if self.branch_checkbox.get():  # need to fix the branch name ~
            if self.branch_column.get() == "":
                self.update_status("Please provide the column for splitting by value", "red")
                return
        else:
            if self.chunk_size.get() == "":
                self.update_status("Please provide the chunk size", "red")
                return

        self.progress_bar.configure(mode="indeterminate", progress_color="lightblue")
        self.progress_bar.start()
        self.update_status("Please wait...", "blue")
        threading.Thread(target=self.generate_files).start()

    @staticmethod
    def remove_extra_spaces(sentence):
        cleaned_sentence = re.sub(r' +', ' ', sentence).strip()
        return cleaned_sentence
    
    def extract_data(self, content):

        format_string = ""
        self.concat_string = ""
        matches = re.finditer(ExcelHyperlinkSplitter.pattern, content, re.MULTILINE)

        for match in matches:
            
            # Extract data from each group
            column = match.group(1)
            others = match.group(2)
            square = match.group(3)

            if column is not None:
                column_index = excel_column_letter_to_index(column)
                try:
                    cell_value = self.df.iloc[0, column_index]

                    if not isinstance(cell_value, str):
                        if pd.isna(cell_value):  # Assuming you are working with pandas NaN
                            cell_value = "N/A"
                        else:
                            cell_value = str(int(cell_value)) if not isinstance(cell_value, datetime) else str(cell_value)

                    self.concat_string += f"IF(ISBLANK({column}#), \"N/A\", {column}#),"
                    format_string += cell_value
                except Exception as e:
                    # print(e)
                    pass

            if others is not None:

                try:
                    if '&' in others:
                        self.concat_string += f"{others.replace('&', '%26')},"
                    elif '\\n' in others:
                        self.concat_string += f"{others.replace('\\n', '%0A')},"
                        format_string += f"{others.replace("\\n", "\n")[1:-1]}"
                    elif '\\t' in others:
                        self.concat_string += f"{others.replace('\\t', '%09')},"
                        format_string += f"{others.replace('\\t', '\t')[1:-1]}"
                    else:
                        self.concat_string += f"{others},"
                        format_string += others[1:-1]

                except Exception as e:
                    pass

            if square is not None:
                square = square.upper()
                mode, index = square.split('.')
                mode = mode[1:]
                column_index = excel_column_letter_to_index(index[:-1])
                self.update_status(f"", "black")

                try:
                    if "DATE" == mode:
                        # self.concat_string += f"TEXT({index[:-1]}#,\"dd-mm-yyyy\"),"
                        self.concat_string += f"IF(ISBLANK({index[:-1]}#), \"N/A\", TEXT({index[:-1]}#,\"dd-mm-yyyy\")),"

                        cell_value = self.df.iloc[0, column_index]

                        if pd.isna(cell_value):
                            format_string += "N/A"
                        else:
                            if isinstance(cell_value, datetime):
                                date_value = cell_value
                            else:
                                if not isinstance(cell_value, str):
                                    cell_value = str(cell_value)

                                serial_number = int(cell_value)
                                date_value = datetime(1899, 12, 30) + timedelta(days=serial_number)

                            format_string += date_value.strftime('%d-%m-%Y')

                    elif "AMPR" == mode:
                        self.concat_string += f"SUBSTITUTE({index[:-1]}#, \"&\", \"%26\"),"

                        cell_value = self.df.iloc[0, column_index]

                        if pd.isna(cell_value):
                            format_string += "N/A"
                        else:
                            if not isinstance(cell_value, str):
                                cell_value = str(cell_value)

                            format_string += cell_value
                    else:
                        self.update_status(f"Invalid mode '{mode}' is being used!", "orange")

                        
                except Exception as e:
                    self.update_status(str(e), "red")
                    
        format_string = ExcelHyperlinkSplitter.remove_extra_spaces(format_string)

        self.preview_frame[1].configure(state="normal")
        self.preview_frame[1].delete("1.0", "end")
        self.preview_frame[1].insert("end", format_string)
        self.preview_frame[1].configure(state="disabled")



    def on_text_change(self, event=None):
        if self.format_frame[1].edit_modified():
            content = self.format_frame[1].get("1.0", "end-1c")
            self.extract_data(content)
            self.format_frame[1].edit_modified(False)
            self.preview_frame[1].yview(ctk.END)

    def create_frame(self, row, height, label, state="normal", margine_y=15):
        # Create a Label for the Entry box and place it at [row, 0]
        label_widget = ctk.CTkLabel(self.root, text=label, font=("Comic Sans MS", 18, "bold"))
        label_widget.grid(row=row, column=0, padx=(15, 5), pady=(margine_y, 5), sticky="w")

        frame_widget = ctk.CTkTextbox(self.root, height=height, state=state, font=("Courier New", 15), wrap="word",  spacing1=5, spacing2=5, spacing3=5)
        frame_widget.grid(row=row+1, column=0, columnspan=4, sticky="ew", padx=(15, 15), pady=(0, 0))
        self.root.columnconfigure(0, weight=1)
        return label_widget, frame_widget

    def create_widgets(self, row, label, type, margine_y=15):
        # Create a Label for the Entry box and place it at [row, 0]
        label_widget = ctk.CTkLabel(self.root, text=label, font=("Comic Sans MS", 18, "bold"))
        label_widget.grid(row=row, column=0, padx=(15, 5), pady=(margine_y, 5), sticky="w")

        # Create an Entry widget and place it at [row, 1] with columnspan=1 to take all remaining space
        entry_var = ctk.StringVar()
        entry_box_widget = ctk.CTkEntry(self.root, textvariable=entry_var, font=("Courier New", 14))
        entry_box_widget.grid(row=row+1, column=0, columnspan=2, sticky="ew", padx=(15, 5), pady=(0, 0), ipadx=5, ipady=5)

       # Create a File Picker button and place it at [row, 3]
        file_picker_button = ctk.CTkButton(self.root, text=f"SELECT", command=lambda: self.open_file_dialog(entry_var, type), height=34, font=("Helvetica", 14, "bold"))
        
        file_picker_button.grid(row=row+1, column=3, sticky="e", padx=(5, 15), pady=(0, 0))
        return label_widget, entry_box_widget, file_picker_button, entry_var

    def open_file_dialog(self, entry_var, type):
        if type == "askopenfilename":
            self.file_path = askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xls")])
            if self.file_path != "":
                entry_var.set(self.file_path)
                self.excel_widgets[2].configure(state="disabled")
                self.progress_bar.configure(mode="indeterminate", progress_color="lightblue")
                self.progress_bar.start()
                self.update_status("please wait...", "blue")
                threading.Thread(target=self.load_excel_file, args=(self.file_path,)).start()
        else:
            self.output_file_path = askdirectory()
            entry_var.set(self.output_file_path)

    def load_excel_file(self, file_path):
        try:
            self.df = pd.read_excel(file_path)

            self.excel_widgets[2].configure(state="normal")

            # Common progress bar configuration
            self.configure_progress_bar("lightgreen")

            # # Enable format_frame and mark it as modified
            self.format_frame[1].configure(state="normal")
            self.format_frame[1].edit_modified(True)

            # Trigger text change event
            self.on_text_change()

            self.update_status("Successfully Loaded!", "green")
            
        except Exception as e:
            self.configure_progress_bar("red")
            self.update_status(str(e), "red")

    def configure_progress_bar(self, color):
        self.progress_bar.configure(mode="determinate", progress_color=color)
        self.progress_bar.stop()
        self.progress_bar.set(1)

    def update_status(self, info, color):
        self.statusVar.set(f"[info]: {info}")
        self.status.configure(text_color=color)


if __name__ == "__main__":
    root = ctk.CTk()
    app = ExcelHyperlinkSplitter(root)
    root.mainloop()

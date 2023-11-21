# Import necessary modules
# tkinter for the GUI, pandas for data processing, matplotlib for plotting, 
# sklearn for stratified sampling, threading for running tasks in the background, 
# and other utilities.
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import matplotlib.pyplot as plt
from sklearn.model_selection import train_test_split
import threading
from Stratified_sample_SK import sample_and_plot_from_excel
import os
import ctypes

# Try to enhance the application's interface on Windows by setting the DPI awareness level.
try:
    PROCESS_SYSTEM_DPI_AWARE = 1
    ctypes.windll.shcore.SetProcessDpiAwareness(PROCESS_SYSTEM_DPI_AWARE)
except (AttributeError, OSError):
    pass  # If it fails, the application will still work, just not DPI-aware.

# Function to open a dialog to select a file path and update an entry widget with the path.
def select_file_path(entry_widget, saving=False):
    current_directory = os.path.dirname(os.path.realpath(__file__))
    if saving:
        file_path = filedialog.asksaveasfilename(initialdir=current_directory, defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx;*.xls"), ("All files", "*.*")])
    else:
        file_path = filedialog.askopenfilename(initialdir=current_directory)
    if file_path:
        entry_widget.delete(0, tk.END)
        entry_widget.insert(0, file_path)

# Function to initiate the data processing task with inputs from the GUI.
def start_processing(input_entry, output_entry, column_entry, sample_entry, sheet_entry, log_text):
    input_file = input_entry.get()
    output_file = output_entry.get()
    strata_column = column_entry.get()
    sheet_name = sheet_entry.get() or None

    try:
        num_samples = int(sample_entry.get())
    except ValueError:
        messagebox.showerror("Error", "Number of samples must be an integer.")
        return

    if not input_file or not output_file or not strata_column:
        messagebox.showerror("Error", "Please complete all required fields.")
        return

    processing_thread = threading.Thread(target=sample_and_plot, args=(input_file, output_file, strata_column, num_samples, sheet_name, log_text))
    processing_thread.start()

# Function to handle the stratified sampling and plotting process, including error handling.
def sample_and_plot(input_file, output_file, strata_column, num_samples, sheet_name, log_text):
    try:
        log_text.insert(tk.END, "Processing started...\n")
        sample_and_plot_from_excel(input_file, output_file, strata_column, num_samples, sheet_name)
        log_text.insert(tk.END, "Processing complete.\n")
    except Exception as e:
        messagebox.showerror("Error", str(e))
        log_text.insert(tk.END, f"Error: {str(e)}\n")

# Function to create a footer in the GUI.
def create_footer(parent, trademark_text, bg_color):
    footer = tk.Frame(parent, bg=bg_color)
    footer.grid(row=2, column=0, sticky="ew")
    parent.grid_columnconfigure(0, weight=1)
    parent.grid_rowconfigure(2, weight=0)
    label = tk.Label(footer, text=trademark_text, bg=bg_color, fg="white")
    label.pack()
    return footer

# Set up the main application window.
root = tk.Tk()
root.title("KAF DB Sampling App")
root.geometry("800x650") 
KAF_COLOR = "#61232e"
root.config(bg=KAF_COLOR)

# Create and configure a frame in the application for widgets.
main_frame = ttk.Frame(root)
main_frame.grid(sticky=(tk.W, tk.E, tk.N, tk.S))

# Create a footer in the application.
footer_frame = create_footer(main_frame, "© 2023 Haiqal @ KAF", KAF_COLOR)
footer_frame.grid(row=2, column=0, sticky=tk.EW)

# Adjust the application's style.
style = ttk.Style(root)
style.theme_use("default")
style.configure("TButton", background=KAF_COLOR, foreground="white", highlightbackground= "#61221e")
style.configure("TLabel", background=KAF_COLOR, foreground="white", highlightbackground= "#61221e")
style.configure("TFrame", background=KAF_COLOR)

# Set up a content frame within the main frame and add widgets for file input, output, and processing parameters.
frame = ttk.Frame(main_frame, padding="12 12 12 12")
frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
main_frame.columnconfigure(0, weight=1)
main_frame.rowconfigure(1, weight=1) 

# Widgets for file inputs, strata column, number of samples, optional sheet name, and a log area.
# ... [code setting up input_entry, output_entry, column_entry, sample_entry, sheet_entry, log_text]

# Button to start the processing.
ttk.Button(frame, text="Start Processing", command=lambda: start_processing(input_entry, output_entry, column_entry, sample_entry, sheet_entry, log_text)).grid(column=2, row=6, sticky=tk.W)

for child in frame.winfo_children():
    child.grid_configure(padx=5, pady=5)

# Start the application's main loop, waiting for user interaction.
root.mainloop()

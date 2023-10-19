import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import matplotlib.pyplot as plt
from sklearn.model_selection import train_test_split
import threading
from Stratified_sample_SK import sample_and_plot_from_excel

# Define the KAF Investment Banking color theme
KAF_COLOR = "#800020"  # A shade of burgundy

def select_file_path(entry_widget, saving=False):
    """Populate the entry widget with the user-selected file path."""
    if saving:
        # This is for saving a file (output), so we use 'asksaveasfilename'
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx;*.xls"), ("All files", "*.*")])
    else:
        # This is for opening a file (input), so we use 'askopenfilename'
        file_path = filedialog.askopenfilename()

    entry_widget.delete(0, tk.END)  # Remove current text
    entry_widget.insert(0, file_path)  # Insert the file path


def start_processing(input_entry, output_entry, column_entry, sample_entry, sheet_entry, log_text):
    """Start the data processing with inputs from the GUI."""
    input_file = input_entry.get()
    output_file = output_entry.get()
    strata_column = column_entry.get()
    sheet_name = sheet_entry.get() or None  # None if left empty

    try:
        num_samples = int(sample_entry.get())
    except ValueError:
        messagebox.showerror("Error", "Number of samples must be an integer.")
        return

    if not input_file or not output_file or not strata_column:
        messagebox.showerror("Error", "Please complete all required fields.")
        return

    # Start a thread to run the task
    processing_thread = threading.Thread(
        target=sample_and_plot,  # Passing the function, not calling it
        args=(input_file, output_file, strata_column, num_samples, sheet_name, log_text)
    )
    processing_thread.start()

def sample_and_plot(input_file, output_file, strata_column, num_samples, sheet_name, log_text):
    """Perform the stratified sampling, plot the results, and handle any errors."""
    try:
        # Log start
        log_text.insert(tk.END, "Processing started...\n")

        # Call the existing function, assuming it's defined correctly in 'Stratified_sample_SK.py'
        sample_and_plot_from_excel(input_file, output_file, strata_column, num_samples, sheet_name)

        # Log completion
        log_text.insert(tk.END, "Processing complete.\n")
    except Exception as e:
        messagebox.showerror("Error", str(e))
        log_text.insert(tk.END, f"Error: {str(e)}\n")

# Set up the main application window
root = tk.Tk()
root.title("KAF Investment Sampling Tool")
root.geometry("700x500")
root.config(bg=KAF_COLOR)

style = ttk.Style(root)
style.theme_use("default")
style.configure("TButton", background=KAF_COLOR, foreground="white")
style.configure("TLabel", background=KAF_COLOR, foreground="white")
style.configure("TFrame", background=KAF_COLOR)

frame = ttk.Frame(root, padding="12 12 12 12")
frame.grid(column=0, row=0, sticky=(tk.W, tk.E, tk.N, tk.S))
frame.columnconfigure(0, weight=1)
frame.rowconfigure(0, weight=1)

# Input file selection
ttk.Label(frame, text="Input File:").grid(column=1, row=1, sticky=tk.W)
input_entry = ttk.Entry(frame, width=40)
input_entry.grid(column=2, row=1, sticky=(tk.W, tk.E))
ttk.Button(frame, text="Browse", command=lambda: select_file_path(input_entry)).grid(column=3, row=1, sticky=tk.W)

# Output file selection
ttk.Label(frame, text="Output File:").grid(column=1, row=2, sticky=tk.W)
output_entry = ttk.Entry(frame, width=40)
output_entry.grid(column=2, row=2, sticky=(tk.W, tk.E))

# Strata column selection
ttk.Label(frame, text="Strata Column:").grid(column=1, row=3, sticky=tk.W)
column_entry = ttk.Entry(frame, width=40)
column_entry.grid(column=2, row=3, sticky=(tk.W, tk.E))

# Number of samples selection
ttk.Label(frame, text="Number of Samples:").grid(column=1, row=4, sticky=tk.W)
sample_entry = ttk.Entry(frame, width=40)
sample_entry.grid(column=2, row=4, sticky=(tk.W, tk.E))

# Sheet name selection (optional)
ttk.Label(frame, text="Sheet Name (optional):").grid(column=1, row=5, sticky=tk.W)
sheet_entry = ttk.Entry(frame, width=40)
sheet_entry.grid(column=2, row=5, sticky=(tk.W, tk.E))

# Log text widget
log_text = tk.Text(frame, height=10, width=50)
log_text.grid(column=1, row=7, columnspan=3, sticky=(tk.W, tk.E))

# Start button
ttk.Button(frame, text="Start Processing", command=lambda: start_processing(input_entry, output_entry, column_entry, sample_entry, sheet_entry, log_text)).grid(column=2, row=6, sticky=tk.W)

for child in frame.winfo_children():
    child.grid_configure(padx=5, pady=5)

root.mainloop()

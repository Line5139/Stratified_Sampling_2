import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import matplotlib.pyplot as plt
from sklearn.model_selection import train_test_split
import threading
from Stratified_sample_SK import sample_and_plot_from_excel
import os

def select_file_path(entry_widget, saving=False):
    """Populate the entry widget with the user-selected file path."""
    # Get the directory of the current script.
    current_directory = os.path.dirname(os.path.realpath(__file__))

    if saving:
        # This is for saving a file (output), so we use 'asksaveasfilename'.
        # We set 'initialdir' to the script's directory.
        file_path = filedialog.asksaveasfilename(
            initialdir=current_directory,
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx;*.xls"), ("All files", "*.*")]
        )
    else:
        # This is for opening a file (input), so we use 'askopenfilename'.
        # We set 'initialdir' to the script's directory.
        file_path = filedialog.askopenfilename(
            initialdir=current_directory
        )

    # If a file path is selected, update the entry widget.
    if file_path:
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


def create_footer(parent, trademark_text, bg_color):
    """Create a footer with a trademark text."""
    footer = tk.Frame(parent, bg=bg_color)  # Set the background color here
    footer.grid(row=2, column=0, sticky="ew")  # The footer expands horizontally (east-west)
    
    parent.grid_columnconfigure(0, weight=1)  # Allows the footer to expand horizontally with the window resizing
    parent.grid_rowconfigure(2, weight=0)  # The footer's height doesn't change with the window resizing
    
    # Place the trademark text
    label = tk.Label(footer, text=trademark_text, bg=bg_color, fg="white")  # Set the text color to white
    label.pack()

    return footer


# Set up the main application window
root = tk.Tk()
root.title("KAF DB Sampling App")
root.geometry("512x530")  # increased height to accommodate header and footer
root.resizable(False, False)

# Define the KAF Investment Banking color theme
KAF_COLOR = "#61232e"
root.config(bg=KAF_COLOR)

# Create a main frame that will contain other widgets and frames
main_frame = ttk.Frame(root)
main_frame.grid(sticky=(tk.W, tk.E, tk.N, tk.S))  # Use grid here

# Create footer
footer_frame = create_footer(main_frame, "© 2023 Haiqal @ KAF", KAF_COLOR)  # Attach the footer to the main frame
footer_frame.grid(row=2, column=0, sticky=tk.EW)  # Place the footer in the grid

style = ttk.Style(root)
style.theme_use("default")
style.configure("TButton", background=KAF_COLOR, foreground="white", highlightbackground= "#61221e")
style.configure("TLabel", background=KAF_COLOR, foreground="white", highlightbackground= "#61221e")
style.configure("TFrame", background=KAF_COLOR)

# Create content frame and add it to the grid of the main frame
frame = ttk.Frame(main_frame, padding="12 12 12 12")
frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

main_frame.columnconfigure(0, weight=1)
main_frame.rowconfigure(1, weight=1) 

# Input file selection
ttk.Label(frame, text="Input File:").grid(column=1, row=1, sticky=tk.W)
input_entry = ttk.Entry(frame, width=40)
input_entry.grid(column=2, row=1, sticky=(tk.W, tk.E))
ttk.Button(frame, text="Browse", command=lambda: select_file_path(input_entry)).grid(column=3, row=1, sticky=tk.W)

# Output file selection
ttk.Label(frame, text="Output File:").grid(column=1, row=2, sticky=tk.W)
output_entry = ttk.Entry(frame, width=40)
output_entry.grid(column=2, row=2, sticky=(tk.W, tk.E))
ttk.Button(frame, text="Browse", command=lambda: select_file_path(output_entry, saving=True)).grid(column=3, row=2, sticky=tk.W)


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

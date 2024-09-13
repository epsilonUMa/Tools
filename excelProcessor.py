import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk

def load_file():
    file_path = filedialog.askopenfilename(title="Select Excel File", filetypes=[("Excel files", "*.xlsx *.xls")])
    if file_path:
        try:
            # Load the Excel file without headers
            global df
            df = pd.read_excel(file_path, header=None)
            messagebox.showinfo("Success", "File loaded successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load the file: {e}")

def process_data():
    if df is not None:
        try:
            # Split the content of the first column (A) by '.' and append the split values
            df_expanded = df[0].str.split('.', expand=True)
            # Concatenate the original column with the split values
            global df_result
            df_result = pd.concat([df, df_expanded], axis=1)
            messagebox.showinfo("Success", "Data processed successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to process the data: {e}")
    else:
        messagebox.showwarning("Warning", "Please load a file first.")

def save_file():
    global output_file_path
    if df_result is not None:
        output_file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if output_file_path:
            try:
                df_result.to_excel(output_file_path, index=False, header=False)
                messagebox.showinfo("Success", "File saved successfully!")
                cal_output_file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
                df_new = pd.read_excel(output_file_path, header=None)

                # Drop the first column (A) and count occurrences of unique elements starting from column B
                element_counts_new = df_new.iloc[:, 1:].stack().value_counts()

                # Display or save the result
                print(element_counts_new)

                # If needed, save the result to a new Excel file
                element_counts_new.to_excel(cal_output_file_path, header=False)
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save the file: {e}")
    else:
        messagebox.showwarning("Warning", "No data to save. Please process the data first.")


# Initialize the GUI application
app = tk.Tk()
app.title("Excel Data Processor")
app.geometry("400x200")

# Global variables for dataframes
df = None
df_result = None

# Create GUI elements
load_button = ttk.Button(app, text="Load Excel File", command=load_file)
load_button.pack(pady=10)

process_button = ttk.Button(app, text="Process Data", command=process_data)
process_button.pack(pady=10)

save_button = ttk.Button(app, text="Save Processed File", command=save_file)
save_button.pack(pady=10)

# calculate_button = ttk.Button(app, text="Calculate times", command=calculate)
# calculate_button.pack(pady=10)

# Start the GUI loop
app.mainloop()

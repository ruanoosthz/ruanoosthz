import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import threading
import sys
import main as rion


# -------------------------------------------------------------------
# Wrapper that calls your original script
# -------------------------------------------------------------------
def process_all_rnd(input_folder, output_file):
    rion.data_folder = input_folder
    rion.output_workbook = output_file
    rion.main()


# -------------------------------------------------------------------
# A class that redirects stdout to a Tkinter text widget
# -------------------------------------------------------------------
class TextRedirector:
    def __init__(self, text_widget, tag="stdout"):
        self.text_widget = text_widget
        self.tag = tag

    def write(self, message):
        self.text_widget.configure(state="normal")
        self.text_widget.insert("end", message, (self.tag,))
        self.text_widget.see("end")
        self.text_widget.configure(state="disabled")

    def flush(self):
        pass  # Required for Python stdout compatibility


# -------------------------------------------------------------------
# GUI class
# -------------------------------------------------------------------
class RNDConverterGUI(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("RND → Excel Converter")
        self.geometry("700x500")
        self.resizable(False, False)

        frame = ttk.Frame(self, padding=20)
        frame.pack(fill="both", expand=True)

        # ---------------- INPUT ----------------
        ttk.Label(frame, text="Input Folder (contains .rnd files):").pack(anchor="w")
        input_frame = ttk.Frame(frame)
        input_frame.pack(fill="x")
        self.input_var = tk.StringVar()
        ttk.Entry(input_frame, textvariable=self.input_var).pack(side="left", fill="x", expand=True)
        ttk.Button(input_frame, text="Browse", command=self.select_input_folder).pack(side="left", padx=5)

        # ---------------- OUTPUT ----------------
        ttk.Label(frame, text="Destination Excel File:").pack(anchor="w", pady=(10, 0))
        output_frame = ttk.Frame(frame)
        output_frame.pack(fill="x")
        self.output_var = tk.StringVar()
        ttk.Entry(output_frame, textvariable=self.output_var).pack(side="left", fill="x", expand=True)
        ttk.Button(output_frame, text="Browse", command=self.select_output_file).pack(side="left", padx=5)

        # ---------------- PROGRESS ----------------
        ttk.Label(frame, text="Progress:").pack(anchor="w", pady=(10, 0))
        self.progress = ttk.Progressbar(frame, mode="indeterminate")
        self.progress.pack(fill="x", pady=(0, 10))

        # ---------------- OUTPUT TEXT ----------------
        ttk.Label(frame, text="Log Output:").pack(anchor="w")
        self.log_box = tk.Text(frame, height=15, state="disabled", wrap="word")
        self.log_box.pack(fill="both", expand=True)
        scrollbar = ttk.Scrollbar(self.log_box, command=self.log_box.yview)
        self.log_box["yscrollcommand"] = scrollbar.set
        scrollbar.pack(side="right", fill="y")

        # Capture stdout
        sys.stdout = TextRedirector(self.log_box, "stdout")
        sys.stderr = TextRedirector(self.log_box, "stderr")

        # ---------------- START BUTTON ----------------
        self.start_btn = ttk.Button(frame, text="Start Conversion", command=self.start_thread)
        self.start_btn.pack(pady=10)

    # -------------------------------------------------
    # File dialogs
    # -------------------------------------------------
    def select_input_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.input_var.set(folder)

    def select_output_file(self):
        file = filedialog.asksaveasfilename(
            title="Save Excel File",
            defaultextension=".xlsx",
            filetypes=[("Excel Workbook", "*.xlsx")]
        )
        if file:
            self.output_var.set(file)

    # -------------------------------------------------
    # Thread wrapper
    # -------------------------------------------------
    def start_thread(self):
        input_folder = self.input_var.get()
        output_file = self.output_var.get()

        if not os.path.isdir(input_folder):
            messagebox.showerror("Error", "Please select a valid input folder.")
            return

        if not output_file.endswith(".xlsx"):
            messagebox.showerror("Error", "Please choose a valid .xlsx file.")
            return

        self.progress.start(10)
        self.start_btn.config(state="disabled")

        thread = threading.Thread(target=self.run_conversion, args=(input_folder, output_file))
        thread.start()

    # -------------------------------------------------
    # Background execution of your script
    # -------------------------------------------------
    def run_conversion(self, input_folder, output_file):
        try:
            print("Starting conversion...\n")
            process_all_rnd(input_folder, output_file)
            print("\n✔ Conversion completed successfully!")

        except Exception as e:
            print(f"\nERROR: {e}")

        finally:
            self.progress.stop()
            self.start_btn.config(state="normal")


# -------------------------------------------------------------------
# Launch GUI
# -------------------------------------------------------------------
if __name__ == "__main__":
    app = RNDConverterGUI()
    app.mainloop()

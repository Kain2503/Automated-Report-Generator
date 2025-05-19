import pandas as pd
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
import tkinter as tk
from tkinter import filedialog, messagebox
import os


# Function to generate the PDF report
def generate_report(file_path):
    try:
        # Step 1: Read the uploaded CSV file
        data = pd.read_csv(file_path)

        # Step 2: Extract column names
        columns = data.columns.tolist()

        # Step 3: Prepare the data for the report
        total_companies = len(data)

        # Create a PDF report
        pdf_filename = "Company_Data_Report.pdf"
        c = canvas.Canvas(pdf_filename, pagesize=letter)

        # Title
        c.setFont("Helvetica-Bold", 16)
        c.drawString(200, 750, "Company Data Report")

        # Adding basic information
        c.setFont("Helvetica", 12)
        c.drawString(50, 720, f"Total Companies: {total_companies}")

        # Create a table with company data dynamically
        c.setFont("Helvetica-Bold", 10)
        y_position = 700

        # Print column headers
        for i, column in enumerate(columns):
            c.drawString(50 + (i * 100), y_position, column)
        y_position -= 20

        # Print company data rows
        c.setFont("Helvetica", 10)
        for i in range(min(20, total_companies)):  # Limit to first 20 companies for readability
            for j, column in enumerate(columns):
                c.drawString(50 + (j * 100), y_position, str(data[column][i]))
            y_position -= 20
            if y_position < 100:  # If we run out of space, create a new page
                c.showPage()
                c.setFont("Helvetica-Bold", 10)
                y_position = 750
                for j, column in enumerate(columns):
                    c.drawString(50 + (j * 100), y_position, column)
                y_position -= 20

        # Save the PDF
        c.showPage()
        c.save()

        messagebox.showinfo("Success", f"Report generated successfully!")
        return pdf_filename  # Return the generated PDF filename

    except Exception as e:
        messagebox.showerror("Error", f"Error generating report: {str(e)}")
        return None


# Function to open file dialog and select a CSV file
def upload_file():
    file_path = filedialog.askopenfilename(title="Select CSV File",
                                           filetypes=(("CSV Files", "*.csv"), ("All Files", "*.*")))
    if file_path:
        # Generate report for the uploaded file
        pdf_filename = generate_report(file_path)
        if pdf_filename:
            enable_download_button(pdf_filename)


# Function to enable the Download button after report generation
def enable_download_button(pdf_filename):
    # Enable the download button after generating the report
    download_button.config(state=tk.NORMAL, command=lambda: download_report(pdf_filename))


# Function to download the report
def download_report(pdf_filename):
    try:
        if os.path.exists(pdf_filename):
            os.startfile(pdf_filename)  # For Windows
            messagebox.showinfo("Download", "Report opened. You can also find it in the current directory.")
        else:
            messagebox.showerror("Error", "Report file not found!")
    except Exception as e:
        messagebox.showerror("Error", f"Error downloading the report: {str(e)}")


# Function to create the main GUI window
def create_gui():
    global download_button  # Declare the button as global to access it in other functions
    window = tk.Tk()
    window.title("Automated Report Generation")

    # Set window size
    window.geometry("400x250")

    # Label to show instructions
    label = tk.Label(window, text="Choose a CSV file to generate a Company Report", font=("Arial", 12))
    label.pack(pady=20)

    # Button to upload CSV file
    upload_button = tk.Button(window, text="Upload CSV File", font=("Arial", 12), command=upload_file)
    upload_button.pack(pady=10)

    # Button to download the report (initially disabled)
    download_button = tk.Button(window, text="Download Report", font=("Arial", 12), state=tk.DISABLED)
    download_button.pack(pady=10)

    # Button to exit the program
    exit_button = tk.Button(window, text="Exit", font=("Arial", 12), command=window.quit)
    exit_button.pack(pady=10)

    # Start the GUI loop
    window.mainloop()


# Run the GUI
if __name__ == "__main__":
    create_gui()

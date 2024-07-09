import tkinter as tk
from tkinter import messagebox
import openpyxl
from datetime import datetime

# Function to mark attendance
def mark_attendance():
    student_id = student_id_entry.get()
    current_date = datetime.now().strftime("%Y-%m-%d")
    current_time = datetime.now().strftime("%H:%M:%S")

    # Check if student ID is provided
    if student_id == "":
        messagebox.showerror("Error", "Please enter Student ID")
        return

    # Open the Excel workbook
    try:
        wb = openpyxl.load_workbook("attendance.xlsx")
        sheet = wb.active
    except FileNotFoundError:
        # If file does not exist, create a new workbook with headers
        wb = openpyxl.Workbook()
        sheet = wb.active
        sheet["A1"] = "Date"
        sheet["B1"] = "Student ID"
        sheet["C1"] = "Time"
        wb.save("attendance.xlsx")
    
    # Check if student ID is already marked for today
    for row in sheet.iter_rows(min_row=2, max_col=2, max_row=sheet.max_row):
        if row[0].value == current_date and row[1].value == student_id:
            messagebox.showwarning("Warning", f"Attendance already marked for Student ID {student_id} today")
            return
    
    # Append new attendance record
    next_row = sheet.max_row + 1
    sheet.cell(row=next_row, column=1, value=current_date)
    sheet.cell(row=next_row, column=2, value=student_id)
    sheet.cell(row=next_row, column=3, value=current_time)
    wb.save("attendance.xlsx")
    
    # Show success message
    messagebox.showinfo("Success", f"Attendance marked for Student ID {student_id}")

# Function to generate attendance report
def generate_report():
    try:
        wb = openpyxl.load_workbook("attendance.xlsx")
        sheet = wb.active

        # Create a new workbook for report
        report_wb = openpyxl.Workbook()
        report_sheet = report_wb.active
        report_sheet["A1"] = "Date"
        report_sheet["B1"] = "Student ID"
        report_sheet["C1"] = "Time"

        # Copy data from attendance sheet to report sheet
        for row in sheet.iter_rows(values_only=True):
            report_sheet.append(row)

        report_wb.save("attendance_report.xlsx")
        messagebox.showinfo("Report Generated", "Attendance report generated successfully")
    
    except FileNotFoundError:
        messagebox.showerror("Error", "Attendance data not found. Please mark attendance first")

# GUI setup
root = tk.Tk()
root.title("Attendance Management System")

# Labels and Entry fields
tk.Label(root, text="Student ID: ").grid(row=0, column=0, padx=10, pady=10)
student_id_entry = tk.Entry(root)
student_id_entry.grid(row=0, column=1, padx=10, pady=10)

# Buttons
mark_button = tk.Button(root, text="Mark Attendance", command=mark_attendance)
mark_button.grid(row=1, column=0, padx=10, pady=10)

report_button = tk.Button(root, text="Generate Report", command=generate_report)
report_button.grid(row=1, column=1, padx=10, pady=10)

# Run the GUI application
root.mainloop()

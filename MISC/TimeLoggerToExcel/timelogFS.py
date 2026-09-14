import tkinter as tk
from tkinter import simpledialog, messagebox
import datetime
import os
import pandas as pd
import re


def fullscreen_prompt(title, prompt, multiline=False, cancel_button=False):
    """
    This function displays a full-screen prompt to the user with either a single-line entry field or a multi-line text box.
    """
    # Create a full-screen window for each prompt
    prompt_window = tk.Toplevel()
    prompt_window.title(title)
    prompt_window.attributes("-fullscreen", True)

    # Create a frame to center the label and entry
    frame = tk.Frame(prompt_window)
    frame.pack(expand=True)

    # Display the prompt message
    label = tk.Label(frame, text=prompt, font=("Arial", 24))
    label.pack(pady=20)

    # Variable to store user input
    user_input = tk.StringVar()

    # Check if multiline input is requested
    if multiline:
        # Create a Text widget for multi-line input
        text_box = tk.Text(frame, font=("Arial", 18), height=10, width=70)
        text_box.pack(pady=10)
        
        # Ensure the Text widget is always focused and ready for typing
        text_box.focus_set()
        text_box.focus_force()
        
        # Function to handle submission for the Text widget
        def submit(event=None):
            user_input.set(text_box.get("1.0", "end-1c"))  # Get content from the Text widget
            prompt_window.destroy()

        # Bind the Ctrl+Enter key to the submit function for multi-line input
        text_box.bind("<Control-Return>", submit)
    else:
        # Create an Entry widget for single-line input
        entry = tk.Entry(frame, textvariable=user_input, font=("Arial", 24), width=50)
        entry.pack(pady=10)
        entry.focus_set()
        entry.focus_force()

        # Function to handle submission for the Entry widget
        def submit(event=None):
            prompt_window.destroy()

        # Bind the Enter key to the submit function for single-line input
        entry.bind("<Return>", submit)

    # Button to submit the input manually (if not using Enter)
    submit_button = tk.Button(frame, text="Submit", command=submit, font=("Arial", 20))
    submit_button.pack(pady=20)

    # Add a cancel button if requested
    if cancel_button:
        def cancel():
            user_input.set("CANCEL")
            prompt_window.destroy()

        cancel_button = tk.Button(frame, text="Cancel", command=cancel, font=("Arial", 20))
        cancel_button.pack(pady=10)

    # Wait until the window is closed
    prompt_window.grab_set()
    prompt_window.wait_window()

    return user_input.get()


def validate_time_format(time_str):
    """
    This function validates that the given time string is in the correct format: hh:mm:ss.
    - hh can be from 00-23
    - mm can be from 00-59
    - ss can be from 00-59
    """
    # Pattern to match hh:mm:ss format with hours limited to 00-23
    pattern = r"^([01][0-9]|2[0-3]):[0-5][0-9]:[0-5][0-9]$"
    match = re.match(pattern, time_str)
    if match:
        return True
    else:
        return False


def main():
    # Create the Tkinter root object and hide the main window
    root = tk.Tk()
    root.withdraw()

    # Validate start time input until correct format is entered
    while True:
        start_time = fullscreen_prompt(title="Work Log", prompt="Enter the start time (e.g., 13:10:00):", cancel_button=True)
        if start_time == "CANCEL":
            root.destroy()
            return
        if validate_time_format(start_time):
            break
        else:
            messagebox.showerror("Invalid Format", "The start time must be in the format hh:mm:ss (e.g., 13:10:00) where hh is 00-23, mm and ss are 00-59.")

    # Prompt the user for other inputs
    company = fullscreen_prompt(title="Work Log", prompt="Enter the company name:")
    project = fullscreen_prompt(title="Work Log", prompt="Enter the project name:")
    user_input = fullscreen_prompt(title="Work Log", prompt="What have you been working on?", multiline=True)

    if user_input:
        # Get current date and time
        now = datetime.datetime.now()
        date_str = now.strftime("%Y-%m-%d")
        time_str = now.strftime("%H:%M:%S")

        # Format the date as day-month-year for the file name
        formatted_date = now.strftime("%d-%m-%y")

        # Prepare the data
        data = {
            'Date': [date_str],
            'Time': [time_str],
            'Start Time': [start_time],
            'Company': [company],
            'Project': [project],
            'Entry': [user_input]
        }

        # Define the filename with date appended at the end
        filename = f'c:\\temp\\work_log_{formatted_date}.xlsx'

        # Check if the file exists
        if os.path.exists(filename):
            # Append to existing Excel file
            df_existing = pd.read_excel(filename)
            df_new = pd.DataFrame(data)
            df_combined = pd.concat([df_existing, df_new], ignore_index=True)
            df_combined.to_excel(filename, index=False)
        else:
            # Create a new Excel file
            df_new = pd.DataFrame(data)
            df_new.to_excel(filename, index=False)

    # Destroy the Tkinter root object
    root.destroy()


if __name__ == '__main__':
    main()
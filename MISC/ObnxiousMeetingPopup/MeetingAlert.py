# Install these two python modules before using this script:
# pip install tkinter
# pip install screeninfo

import tkinter as tk
from screeninfo import get_monitors
import sys

def show_fullscreen_meeting_alert(meeting_name):
    root = tk.Tk()
    root.withdraw()  # Hide the root window since we'll use Toplevel windows

    windows = []

    # Create a window on each monitor
    for monitor in get_monitors():
        window = tk.Toplevel(root)
        window.title(f"Meeting Alert - {monitor.name}")
        window.overrideredirect(True)  # Remove window decorations
        window.geometry(f"{monitor.width}x{monitor.height}+{monitor.x}+{monitor.y}")
        window.lift()
        window.attributes("-topmost", True)
        window.focus_force()

        # Set background color to red
        window.configure(bg="red")

        # Create a frame to center the content
        frame = tk.Frame(window, bg="red")
        frame.pack(expand=True)

        # Create the "MEETING ALERT" label
        alert_label = tk.Label(
            frame,
            text="MEETING ALERT",
            font=("Helvetica", 100, "bold"),
            fg="white",
            bg="red"
        )
        alert_label.pack(pady=(0, 50))  # Add some space below the label

        # Create the meeting name label in a slightly smaller font
        meeting_label = tk.Label(
            frame,
            text=meeting_name,
            font=("Helvetica", 70),
            fg="white",
            bg="red",
            wraplength=monitor.width - 100,
            justify="center"
        )
        meeting_label.pack()

        windows.append(window)

    # Create a close button on the primary monitor's window
    primary_window = windows[0] if windows else None
    if primary_window:
        cancel_button = tk.Button(
            primary_window,
            text="Close",
            font=("Helvetica", 20),
            command=root.destroy
        )
        cancel_button.place(
            relx=1.0,
            rely=1.0,
            anchor="se",
            x=-20,
            y=-20
        )

    root.mainloop()

if __name__ == "__main__":
    if len(sys.argv) > 1:
        meeting_name = sys.argv[1]
    else:
        meeting_name = "No Meeting Name Provided"

    show_fullscreen_meeting_alert(meeting_name)

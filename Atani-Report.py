import webbrowser

from tkinter import END, filedialog
import customtkinter
from collect_info import final_report, raw_data

customtkinter.set_appearance_mode("dark")  # Modes: "System" (standard), "Dark", "Light"
customtkinter.set_default_color_theme("blue")  # Themes: "blue" (standard), "green", "dark-blue"

app = customtkinter.CTk()

app_width = 500
app_height = 830


def align_center(width, height):
    """
        Calculates the coordinates to place a window in the center of the screen.

        Args:
            width (int): width of the screen.
            height (int): height of the screen.

        Returns:
            str: A string containing the coordinates to place the window in the center of the screen.
        """

    screen_width = app.winfo_screenwidth()
    screen_height = app.winfo_screenheight()

    x = (screen_width / 2) - (width / 2)
    y = (screen_height / 2) - (height / 2)
    return '%dx%d+%d+%d' % (width, height, x, y)


def display_notification(label, text, disappear_after_ms=0):
    """
        Displays a notification message on a label widget with the given text.
        If `disappear_after_ms` is not 0, the label text will be cleared after the specified time in milliseconds.

        Args:

        - label: A label widget to display the notification on.

        - text: The text to display in the notification.

        - disappear_after_ms: Optional. The time in milliseconds after which the notification should disappear. If not provided or set to 0, the notification will not disappear.
    """
    label.configure(text=text)
    if disappear_after_ms != 0:
        app.after(disappear_after_ms, lambda: label.configure(text=""))


def file_upload():
    print("Upload Button clicked")
    file = filedialog.askopenfilename(title="Select a File",  # initialdir=os.getcwd(),
                                      filetypes=(("Excel files", "*.xlsx*"),
                                                 ("all files", "*.*")))

    # Change entry contents
    file_path_entry.delete(0, END)
    file_path_entry.insert(0, file)
    generate_report_button.configure(state="normal")


def generate_report():
    log_textbox.configure(state="normal")
    log_textbox.delete("1.0", END)
    log_textbox.insert("0.0", "Generating report. Please wait...")

    # Show ERROR if the sheet is not found
    try:
        log_textbox.configure(text_color="white")
        log_textbox.insert("0.0", raw_data(file_path_entry.get()))  # ordered_items_text
    except (KeyError, IndexError):
        log_textbox.configure(text_color="#D0312D")  # red
        log_textbox.insert("0.0", "ERROR! Please check the file or the format is correct.")

    log_textbox.configure(state="disabled")

    report_textbox.delete("1.0", END)
    report_textbox.insert("0.0", final_report(file_path_entry.get()))

    print("Report generated from", file_path_entry.get())


def copy_report():
    app.clipboard_clear()
    app.clipboard_append(report_textbox.get("1.0", END))
    display_notification(label=report_copied_label, text="Copied!", disappear_after_ms=2000)
    print("Report copied to clipboard")


def url(link):
    webbrowser.open_new(link)


app.geometry(align_center(width=app_width, height=app_height))
app.title("Atani Report")

frame_1 = customtkinter.CTkFrame(master=app)
frame_1.pack(pady=20, padx=30, fill="both", expand=True)

header = customtkinter.CTkLabel(master=frame_1,
                                justify=customtkinter.LEFT,
                                text="Atani Report",
                                font=("Roboto", 32))
header.pack(pady=10, padx=10)

# Upload file input and button
file_path_entry = customtkinter.CTkEntry(master=frame_1, placeholder_text="File path")
file_path_entry.pack(pady=10, padx=10, fill="x")
upload_button = customtkinter.CTkButton(master=frame_1,
                                        text="File",
                                        command=file_upload,
                                        width=60,
                                        font=("Roboto", 16))
upload_button.place(relx=1, x=-70, y=68)

# Generate report button
generate_report_button = customtkinter.CTkButton(master=frame_1,
                                                 text="Generate Report",
                                                 command=generate_report,
                                                 width=60,
                                                 font=("Roboto", 16),
                                                 fg_color="transparent",
                                                 border_width=2,
                                                 text_color=("gray10", "#DCE4EE"))
generate_report_button.configure(state="disabled")
generate_report_button.pack(pady=10, padx=10)

# empty line to reduce visual clutter
empty_label_1 = customtkinter.CTkLabel(master=frame_1, text="")
empty_label_1.pack(pady=10, padx=10)

# Log
log_label = customtkinter.CTkLabel(master=frame_1, text="Raw Data", font=("Roboto", 18))
log_label.pack(pady=0, padx=10)

log_textbox = customtkinter.CTkTextbox(master=frame_1,
                                       width=250,
                                       height=100,
                                       fg_color="transparent",
                                       border_width=1,
                                       font=("Roboto", 14))
log_textbox.configure(state="disabled")
log_textbox.pack(pady=10, padx=10, fill="both", expand=True)

# empty line to reduce visual clutter
empty_label_2 = customtkinter.CTkLabel(master=frame_1, text="")
empty_label_2.pack(pady=10, padx=10)

# Report
report_label = customtkinter.CTkLabel(master=frame_1, text="Report", font=("Roboto", 18))
report_label.pack(pady=0, padx=10)

report_textbox = customtkinter.CTkTextbox(master=frame_1, width=250, font=("Roboto", 14))
report_textbox.pack(pady=10, padx=10, fill="both", expand=True)

# "Copied!" text label
report_copied_label = customtkinter.CTkLabel(
    master=frame_1,
    text="",
    font=("Roboto", 16, "bold"),
)
report_copied_label.configure(text_color="#0b8556")  # green
report_copied_label.pack(pady=0)

# Copy button
copy_button = customtkinter.CTkButton(master=frame_1,
                                      text="Copy to clipboard",
                                      command=copy_report,
                                      width=60,
                                      font=("Roboto", 18),
                                      )
copy_button.pack(pady=10, padx=10)

# Footer
footer = customtkinter.CTkLabel(master=frame_1,
                                justify=customtkinter.RIGHT,
                                text="Created by igorsimb.ru",
                                text_color="lightblue",
                                )
footer.pack(padx=10, side="right")
footer.bind("<Button-1>", lambda e: url("https://igorsimb.ru/"))  # <Button-1> = left mouse button

app.mainloop()

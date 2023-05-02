from tkinter import END, messagebox, filedialog

import customtkinter
import webbrowser

from collect_info import final_report, raw_data

customtkinter.set_appearance_mode("dark")  # Modes: "System" (standard), "Dark", "Light"
customtkinter.set_default_color_theme("blue")  # Themes: "blue" (standard), "green", "dark-blue"

app = customtkinter.CTk()


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

    # Show ERROR if the sheet is not found
    try:
        log_textbox.configure(text_color="white")
        log_textbox.insert("0.0", raw_data(file_path_entry.get()))  # ordered_items_text
    except KeyError:
        log_textbox.configure(text_color="#D0312D")  # red
        log_textbox.insert("0.0", "ERROR! Please check the file or the format is correct.")

    log_textbox.configure(state="disabled")

    report_textbox.delete("1.0", END)
    report_textbox.insert("0.0", final_report(file_path_entry.get()))

    print("Report generated", file_path_entry.get())


def copy_report():
    app.clipboard_clear()
    app.clipboard_append(report_textbox.get("1.0", END))
    messagebox.showinfo("Atani Report", "Your report has been copied to clipboard.")
    print("Report copied to clipboard")


def url(link):
    webbrowser.open_new(link)


app.geometry(align_center(width=500, height=830))
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

# Copy button
copy_button = customtkinter.CTkButton(master=frame_1, text="Copy to clipboard", command=copy_report,
                                      width=60, font=("Roboto", 16))
copy_button.pack(pady=10, padx=10)

# Footer
footer = customtkinter.CTkLabel(master=frame_1, justify=customtkinter.RIGHT, text="Created by igorsimb.ru",
                                text_color="lightblue")
footer.pack(padx=10, side="right")
footer.bind("<Button-1>", lambda e: url("https://igorsimb.ru/"))  # <Button-1> = left mouse button

app.mainloop()

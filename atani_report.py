import webbrowser

from tkinter import END, filedialog
import customtkinter as ctk

import collect_info
from collect_info import final_report, get_cell_by_value, raw_data

ctk.set_appearance_mode("dark")  # Modes: "System" (standard), "Dark", "Light"
ctk.set_default_color_theme("blue")  # Themes: "blue" (standard), "green", "dark-blue"


# Functions

def align_center(width, height):
    screen_width, screen_height = app.winfo_screenwidth(), app.winfo_screenheight()
    x, y = (screen_width - width) / 2, (screen_height - height) / 2
    return f"{width}x{height}+{int(x)}+{int(y)}"


def check_correct_file_is_loaded(file):

    """
    Checks if the correct file has been loaded for the selected marketplace. It takes a file as input and tries to
    read data from it. If an error occurs, it displays an error message.
    Otherwise, it checks if the file contains the expected data for the selected marketplace and displays a warning
    message if it doesn't.
    """

    try:
        log_textbox.configure(text_color="white")
        log_textbox.delete("1.0", END)
        log_textbox.insert("0.0", raw_data(file_path_entry.get()))
    except (KeyError, IndexError, TypeError, ValueError):
        report_textbox.delete("1.0", END)
        log_textbox.configure(text_color="#D0312D")  # red
        log_textbox.insert(
            "0.0",
            f"WrongFile ERROR! Could not read file.\nYou picked {collect_info.current_marketplace.upper()} as your "
            f"marketplace. Please check the file or the format is correct.\n\n",
        )
    else:
        ordered_items_cell = get_cell_by_value(
            file_path=file, cell_value="Заказано товаров, шт."
        )
        is_ozon_chosen, is_wb_chosen, is_yandex_chosen = (
            collect_info.current_marketplace == "ozon",
            collect_info.current_marketplace == "wb",
            collect_info.current_marketplace == "yandex",
        )
        # Ozon has "Заказано товаров, шт." in A5, WB in A3, yandex in A4
        if not (
                is_ozon_chosen and "A5" in ordered_items_cell
                or is_wb_chosen and "A3" in ordered_items_cell
                or is_yandex_chosen and "A4" in ordered_items_cell
        ):
            log_textbox.configure(text_color="#d1c413")  # yellow
            log_textbox.insert(
                "0.0",
                f"WrongFile WARNING! You picked {collect_info.current_marketplace.upper()} as your marketplace.\n"
                f"However, the file you uploaded does not contain the expected data for this marketplace.\n "
                f"Please make sure your chosen marketplace matches the uploaded file.\n\n",
            )


def display_notification(label, text, disappear_after_ms=0):
    label.configure(text=text)
    if disappear_after_ms:
        app.after(disappear_after_ms, lambda: label.configure(text=""))


def file_upload():
    file = filedialog.askopenfilename(
        title="Select a File",
        filetypes=[("Excel files", "*.xlsx*"), ("all files", "*.*")],
    )

    file_path_entry.delete(0, END)
    file_path_entry.insert(0, file)


def generate_report():
    """
        Generates a report based on the currently selected marketplace and the file uploaded.
        Displays the report in the report_textbox widget.
    """
    log_textbox.configure(state="normal")
    log_textbox.delete("1.0", END)
    log_textbox.insert("0.0", "Generating report. Please wait...")

    check_correct_file_is_loaded(file_path_entry.get())

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


def choose_marketplace(marketplace):
    """
    Sets the current marketplace to the one passed in the parameter.
    Sets the button colors accordingly.
    Sets the current marketplace in the collect_info module.
    Enables the generate_report_button.
    """
    if marketplace == "ozon":
        ozon_button.configure(fg_color="#025BFB")
        wb_button.configure(fg_color="transparent")
        yandex_button.configure(fg_color="transparent")
        collect_info.current_marketplace = "ozon"
    elif marketplace == "wb":
        wb_button.configure(fg_color="#5B117B")
        ozon_button.configure(fg_color="transparent")
        yandex_button.configure(fg_color="transparent")
        collect_info.current_marketplace = "wb"
    elif marketplace == "yandex":
        yandex_button.configure(fg_color="#FC3F1D")
        wb_button.configure(fg_color="transparent")
        ozon_button.configure(fg_color="transparent")
        collect_info.current_marketplace = "yandex"

    generate_report_button.configure(state="normal")
    return collect_info.current_marketplace


# App

app = ctk.CTk()
app.title("Atani Report")
app_width = 500
app_height = 890
app.geometry(align_center(width=app_width, height=app_height))

frame = ctk.CTkScrollableFrame(app, border_width=1)
frame.pack(pady=20, padx=30, fill="both", expand=True)

# Header label
header = ctk.CTkLabel(frame, text="Atani Report", font=("Roboto", 32), justify=ctk.LEFT)
header.pack(pady=10, padx=10)

# File upload input and button
file_path_entry = ctk.CTkEntry(frame, placeholder_text="File path")
file_path_entry.pack(pady=10, padx=10, fill="x")
upload_button = ctk.CTkButton(frame, text="File", command=file_upload, width=60, font=("Roboto", 16))
upload_button.place(relx=1, x=-70, y=68)

# Marketplace selection buttons
marketplaces_frame = ctk.CTkFrame(frame, fg_color="transparent")
marketplaces_frame.pack(side="top", anchor="w", padx=20, pady=10)

ozon_button = ctk.CTkButton(marketplaces_frame, text="Ozon", font=("Roboto", 24),
                            fg_color="transparent", border_width=1, border_spacing=0,
                            text_color=("gray10", "#DCE4EE"), hover_color="#025BFB",
                            width=80, height=60)
ozon_button.pack(side="left")
ozon_button.bind("<Button-1>", lambda e: choose_marketplace("ozon"))

wb_button = ctk.CTkButton(marketplaces_frame, text="WB", font=("Roboto", 24),
                          fg_color="transparent", border_width=1, border_spacing=0,
                          text_color=("gray10", "#DCE4EE"), hover_color="#5B117B",
                          width=80, height=60)
wb_button.pack(side="left", padx=20)
wb_button.bind("<Button-1>", lambda e: choose_marketplace("wb"))

yandex_button = ctk.CTkButton(marketplaces_frame,
                              text="ЯМ", font=("Roboto", 24),
                              fg_color="transparent", border_width=1,
                              border_spacing=0, text_color=("gray10", "#DCE4EE"),
                              hover_color="#FC3F1D", width=80, height=60)
yandex_button.pack(side="left", padx=0)
yandex_button.bind("<Button-1>", lambda e: choose_marketplace("yandex"))

# Report generation button
generate_report_button = ctk.CTkButton(frame, text="Generate Report", command=generate_report,
                                       width=60, font=("Roboto", 16),
                                       fg_color="transparent", border_width=2,
                                       text_color=("gray10", "#DCE4EE"))
generate_report_button.configure(state="disabled")
generate_report_button.pack(pady=10, padx=10)

# Empty label to reduce visual clutter
empty_label = ctk.CTkLabel(frame, text="")
empty_label.pack(pady=10, padx=10)

# Log label and textbox
log_label = ctk.CTkLabel(frame, text="Raw Data", font=("Roboto", 18))
log_label.pack(pady=0, padx=10)

log_textbox = ctk.CTkTextbox(frame, width=250, height=100, fg_color="transparent",
                             border_width=1, font=("Roboto", 14))
log_textbox.configure(state="disabled")
log_textbox.pack(pady=10, padx=10, fill="both", expand=True)

# Empty label to reduce visual clutter
empty_label_2 = ctk.CTkLabel(master=frame, text="")
empty_label_2.pack(pady=10, padx=10)

# Report label and textbox
report_label = ctk.CTkLabel(master=frame, text="Report", font=("Roboto", 18))
report_label.pack(pady=0, padx=10)

report_textbox = ctk.CTkTextbox(master=frame, width=250, font=("Roboto", 14))
report_textbox.pack(pady=10, padx=10, fill="both", expand=True)

# "Copied!" label
report_copied_label = ctk.CTkLabel(master=frame, text="", font=("Roboto", 16, "bold"))
report_copied_label.configure(text_color="#0b8556")  # green
report_copied_label.pack(pady=0)

# Copy button
copy_button = ctk.CTkButton(master=frame, text="Copy to clipboard", command=copy_report, width=60, font=("Roboto", 18))
copy_button.pack(pady=10, padx=10)

# Footer
footer = ctk.CTkLabel(master=frame, justify=ctk.RIGHT, text="Created by igorsimb.ru", text_color="lightblue")
footer.pack(padx=10, pady=2, side="right")
footer.bind("<Button-1>", lambda e: url("https://igorsimb.ru/"))  # <Button-1> = left mouse button

app.mainloop()

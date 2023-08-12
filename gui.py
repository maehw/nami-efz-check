from bescheinigungs_check import BescheinigungsCheck
import os
from tkinter import TOP, BOTTOM, Label, font
from tkinterdnd2 import *


window_width = 550
window_height = 350
title = 'SGB VIII Check'
color_failed = '#FF681E'
color_success = '#069C56'
color_error = '#FF980E'


def get_path(event):
    full_file_name = event.data
    file_name = os.path.basename(full_file_name)
    print(full_file_name)
    path_label.configure(text=file_name)
    status_label.configure(text='')
    status_label.configure(text='Wird geprüft.')

    check = BescheinigungsCheck.from_pdf_file(full_file_name)
    efz_status = check.get_check_status()

    number = check.get_id()
    firstnames = check.get_firstnames()
    surname = check.get_surname()
    birthdate = check.get_birthdate()

    details_label.configure(text=f"{number} / {firstnames} {surname}, geb. am {birthdate}")

    if efz_status:
        status_label.configure(text='Gültig.')
        root.configure(background=color_success)
    else:
        status_label.configure(text='Ungültig oder Prüfung fehlgeschlagen.')
        root.configure(background=color_failed)


root = TkinterDnD.Tk()
root.geometry(f"{window_width}x{window_height}")
root.title(title)

path_label = Label(root, text="Drag and drop file in the entry box")
path_label.pack(side=TOP, padx=10, pady=10)

status_label = Label(root, text="", padx=5, pady=10)
status_label.pack(side=BOTTOM)

details_label = Label(root, text="", padx=20, pady=20, font=font.Font(weight="bold"))
details_label.pack(side=BOTTOM)

root.drop_target_register(DND_ALL)
root.dnd_bind("<<Drop>>", get_path)

root.mainloop()

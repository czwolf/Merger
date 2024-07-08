# Při převodu na xlsx je potřeba v nástroji auto-py-to-exe v Advanced vyplnit do --hidden-import  hodnotu xlrd
from Merger import Merger
from tkinter import *
import os
from tkinter import filedialog
import glob
import pandas as pd

# path = os.getcwd() + '\spojeno'
# file = Merger(file_type="xlsx", output_filename="spojený soubor xlsx", output_folder=path)
# print(file.merge())

def open_folder():
    folder = load_entry.get()
    try:
        if folder:
            os.startfile(folder)
        else:
            count_label["text"] = "Žádná složka se soubory není vybraná.\n\n"
    except:
        count_label["text"] = "Žádná složka se soubory není vybraná.\n\n"

def create_checkbutton():
    checkbutton.pack()

def hide_checkbutton():
    checkbutton.pack_forget()

def setOutputFileExtension(*args):
    if file_extension.get() == 'csv':
        count_label["text"] = "Hodnoty musí být odděleny středníkem ' ; '\n\nKódování nastaveno na utf-8"
        hide_checkbutton()
        load_entry.delete(0, END)
        name_entry.delete(0, END)
        extension_label["text"] = file_extension.get()
    elif file_extension.get() == 'pdf':
        count_label["text"] = "Primárně určeno ke spojování formátu A4"
        hide_checkbutton()
        load_entry.delete(0, END)
        name_entry.delete(0, END)
        extension_label["text"] = file_extension.get()
    else:
        count_label["text"] = ""
        hide_checkbutton()
        load_entry.delete(0, END)
        name_entry.delete(0, END)
        extension_label["text"] = file_extension.get()

def merge():
    folder = load_entry.get()
    file_name = name_entry.get()
    filenames = glob.glob(folder + "\*"+file_extension.get())
    cnt = len(filenames)
    if file_name:
        if cnt > 0:
            if file_extension.get() == 'pdf':
                pdf = Merger(file_type="pdf", output_filename=file_name, output_folder=load_entry.get(), work_folder=load_entry.get())
                pdf.merge()
                count_label["text"] = ""
                count_label["text"] = f"Soubory úspěšně sloučeny.\n\nNázev souboru: {file_name}.{file_extension.get()}"
            elif file_extension.get() == 'xlsx' or file_extension.get() == 'xls':
                if check.get() == 0:
                    xlsx = Merger(file_type="xlsx", output_filename=file_name, output_folder=load_entry.get(), work_folder=load_entry.get())
                    xlsx.merge(duplicity_keep=True)
                    count_label["text"] = f"Soubory úspěšně sloučeny.\n\nNázev souboru: {file_name}.{file_extension.get()}"
                else:
                    xlsx = Merger(file_type="xlsx", output_filename=file_name, output_folder=load_entry.get(), work_folder=load_entry.get())
                    xlsx.merge(duplicity_keep=False)
                    count_label["text"] = ""
                    count_label["text"] = f"Soubory úspěšně sloučeny.\n\nDuplicity odebrány!\n\nNázev souboru: {file_name}.xlsx"
            elif file_extension.get() == 'csv':
                if check.get() == 0:
                    csv = Merger(file_type="csv", output_filename=file_name, output_folder=load_entry.get(), work_folder=load_entry.get())
                    csv.merge(duplicity_keep=True)
                    count_label["text"] = f"Soubory úspěšně sloučeny.\n\nNázev souboru: {file_name}.{file_extension.get()}"
                else:
                    csv = Merger(file_type="csv", output_filename=file_name, output_folder=load_entry.get(), work_folder=load_entry.get())
                    csv.merge(duplicity_keep=False)
                    count_label["text"] = f"Soubory úspěšně sloučeny.\n\nNázev souboru: {file_name}.{file_extension.get()}"

        else:
            count_label["text"] = f"Není vybraná cesta k souborům, nebo složka neobsahuje soubory {file_extension.get()}.\n\n"
    else:
        count_label["text"] = "Není zadán název výstupního souboru.\n\n"

# def help():
#     w2 = Tk()
#     icon = resource_path("help.ico")
#     w2.minsize(600,600)
#     w2.iconbitmap(icon)
#     w2.resizable(False,False)
#     w2.title("Nápověda")
#
#     frame_title = Frame(w2)
#     frame_title.pack()
#     frame_text = Frame(w2)
#     frame_text.pack()
#
#     title = Label(frame_title, text="Nápověda")
#     title.grid(row=0,column=0, pady=5)
#     text = Text(frame_text, width=110, height=34)
#     with open('postup.txt', "r", encoding="utf8") as file:
#         for row in file:
#             text.insert(INSERT,row)
#     text["state"] = DISABLED
#     text.grid(row=1,column=0, padx=10, pady=5, sticky=E+W+N+S)
#
#     scrollbar = Scrollbar(frame_text)
#     scrollbar.grid(row=1,column=1,sticky=N+S)
#     scrollbar.config(command=text.yview)
#
#     w2.mainloop()

def select_folder():
    try:
        load_entry.delete(0, END)
        checkbutton.deselect()
        info_duplicity["text"] = ""
        source_path = filedialog.askdirectory(title='Výběr adresáře se zdrojovými soubory')
        load_entry.insert(0, source_path)
        folder = load_entry.get()
        filenames = glob.glob(folder + "\*."+file_extension.get())
        cnt = len(filenames)
        count_label["text"] = f"Počet nalezených souborů: {cnt}\n\n"

        if cnt > 0 and (file_extension.get() == 'xlsx' or file_extension.get() == 'xls'):
            all_dfs = pd.DataFrame()
            for file in filenames:
                df = pd.read_excel(file, engine="openpyxl")
                all_dfs = pd.concat([all_dfs, df], ignore_index=True, sort=False)
            duplicate_rows = all_dfs[all_dfs.duplicated()]
            duplicity_count = len(duplicate_rows)
            count_label["text"] = f"Počet nalezených souborů: {cnt}\n\nPočet nalezených duplicitních řádků: {duplicity_count}"
            if duplicity_count > 0:
                create_checkbutton()
            else:
                hide_checkbutton()

        elif cnt > 0 and (file_extension.get() == 'csv'):
            all_dfs = pd.DataFrame()
            for file in filenames:
                df = pd.read_csv(file, sep=';', encoding='utf8')
                all_dfs = all_dfs._append(df, ignore_index=True)
            duplicate_rows = all_dfs[all_dfs.duplicated()]
            duplicity_count = len(duplicate_rows)
            count_label["text"] = f"Počet nalezených souborů: {cnt}\n\nPočet nalezených duplicitních řádků: {duplicity_count}"
            if duplicity_count > 0:
                create_checkbutton()
            else:
                hide_checkbutton()
    except:
        count_label["text"] = 0

if __name__ == '__main__':
    win = Tk()
    win.title("Spojovač souborů PDF, XLS, XLSX, CSV")
    win.minsize(600, 160)
    # win.iconbitmap("merger.ico")
    win.resizable(True, False)
    main_font = ("Sans Serif", 11)

    # definice rámů
    extension_frame = Frame(win)
    extension_frame.pack()
    load_frame = Frame(win)
    load_frame.pack()
    file_name_frame = Frame(win)
    file_name_frame.pack()
    output_frame = Frame(win)
    output_frame.pack()
    duplicity_frame = Frame(win)
    duplicity_frame.pack()
    button_frame = Frame(win)
    button_frame.pack()

    # výběr typu souboru
    file_extension = StringVar(extension_frame, "xlsx")
    file_extension.trace('w', setOutputFileExtension)
    file_extension_title = Label(extension_frame, text="Typ souboru: ", font=main_font)
    file_extension_title.grid(row=0, column=0, padx=5, pady=5)
    xlsx = Radiobutton(extension_frame, text="XLSX, XLS", variable=file_extension, value="xlsx", font=main_font)
    xlsx.grid(row=0, column=1, padx=5, pady=5)
    csv = Radiobutton(extension_frame, text="CSV", variable=file_extension, value="csv", font=main_font)
    csv.grid(row=0, column=2, padx=5, pady=5)
    pdf = Radiobutton(extension_frame, text="PDF", variable=file_extension, value="pdf", font=main_font)
    pdf.grid(row=0, column=3, padx=5, pady=5)

    # načítací část
    load_title = Label(load_frame, text="Cesta k souborům", font=main_font)
    load_title.grid(row=0, column=0, padx=5, pady=5)
    load_entry = Entry(load_frame, width=50, font=main_font)
    load_entry.grid(row=0, column=1, padx=5, pady=5)
    browse_button = Button(load_frame, text="Browse", font=main_font, command=select_folder)
    browse_button.grid(row=0, column=2, padx=5, pady=5)

    # název souboru
    name_title = Label(file_name_frame, text="Název výstupního souboru", font=main_font)
    name_title.grid(row=0, column=0, padx=5, pady=5)
    name_entry = Entry(file_name_frame, width=30, font=main_font)
    name_entry.grid(row=0, column=1, padx=5, pady=5)
    dot_label = Label(file_name_frame, text=".", font=main_font)
    dot_label.grid(row=0, column=2, padx=0, pady=5)
    extension_label = Label(file_name_frame, text=file_extension.get(), font=main_font)
    extension_label.grid(row=0, column=3, padx=5, pady=5)

    # výstupní část
    count_label = Label(output_frame, text="\n\n", font=main_font)
    count_label.grid(row=1, column=1, padx=5, pady=5)
    info_duplicity = Label(output_frame, text="", font=main_font)
    info_duplicity.grid(row=2, column=1, padx=5, pady=5)

    # duplicity část
    check = IntVar()
    checkbutton = Checkbutton(duplicity_frame, text="Odebrat duplicity", onvalue=1, offvalue=0, variable=check, font=main_font)
    checkbutton.pack_forget()

    # button část
    submit_button = Button(button_frame, text="Proveď sloučení", font=main_font, bg="lightgray", command=merge)
    submit_button.grid(row=2, column=0, padx=5, pady=5, ipady=2, ipadx=3)
    folder_button = Button(button_frame, text="Otevřít adresář", font=main_font, bg="lightgray", command=open_folder)
    folder_button.grid(row=2, column=1, padx=5, pady=5, ipady=2, ipadx=3)
    # help_button = Button(button_frame, text="Info", font=main_font, bg="lightgray", command=help)
    # help_button.grid(row=2, column=2, padx=5, pady=5, ipady=2, ipadx=3)

    win.mainloop()
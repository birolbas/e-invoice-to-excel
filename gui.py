from tkinter import *
from tkinter import filedialog
from tkinter import filedialog, messagebox
from tkinter.simpledialog import askstring

import os 
from converter import EInvoiceConverter
 

window = Tk()
window.title("E-Fatura Çevirici")
window.geometry("500x500")

title_label = Label(window, text="E-Faturaları Tabloya Çevirici", font=("Arial", 16, "bold"))
title_label.pack(pady=30)
global converter

def select_folder():
    global pdf_paths, df, converter, folder_path
    folder_path = filedialog.askdirectory()
    pdf_paths = []
    if folder_path:
        for file in os.listdir(folder_path):
            if file.lower().endswith(".pdf"):
                pdf_paths.append(os.path.join(folder_path, file))

    if pdf_paths:
        selected_label.config(text=f"{len(pdf_paths)} fatura seçildi")
        global converter
        converter = EInvoiceConverter(pdf_paths=pdf_paths)    
    else:
        selected_label.config(text="PDF dosyası bulunamadı.")
        converter = EInvoiceConverter(pdf_paths=pdf_paths)    

def create_excel():
    global df 
    df = converter.set_structure()
    if df is not None:
        excel_name = askstring("filename", "Oluşturmak istediğiniz excel dosyasına isim veriniz.")
        converter.create_excel(dataframe=df ,excel_name=excel_name,)
        messagebox.showinfo("Başarılı", "Excel dosyası oluşturuldu!")
    else:
        messagebox.showwarning("Uyarı", "Lütfen önce klasör seçin.")    

folder_button = Button(window, text="Faturaların olduğu klasörü seçiniz.", width=25, command=select_folder)
folder_button.pack(pady=20)

selected_label = Label(window, text="Henüz dosya seçilmedi..", fg="gray")
selected_label.pack(pady=5)

create_excel_button = Button(window, text="Tabloya çevir ", width=25, command=create_excel)
create_excel_button.pack(pady=20)


window.mainloop()
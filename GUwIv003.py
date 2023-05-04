import tkinter as tk
import os
import tkinter.messagebox
from tkinter.filedialog import askopenfilename, asksaveasfilename
from tkinter import ttk
import win32com.client as win32



def cfgfile():
    sciezka = os.listdir()
    plikcfg = "cfg.ini"
    if os.path.exists(plikcfg):
        for x in sciezka:  # sprawdza czy plik cfg.ini istnieje i czy w pliku znajduje sie poprawna sciekz
            if plikcfg in x:
                with open(plikcfg, 'r', encoding="utf-8") as plik:
                    y = plik.read()
                    start_sciezki = y.find("//")
                    obecnasciezk = os.getcwd()  # zmienna obecnej sciezki

                    if y.find("//") == 0:  # sprawdza czy sciezka zaczyna się od odpowiedniego rozwiniecia i zwraca
                        z = y[start_sciezki:]
                        if y.endswith("/"):
                            return z
                        else:
                            return z + '/'
                    elif y.find(':/') == 1:
                        z = y[y.find(':/')-1:]
                        if y.endswith("/"):
                            return z
                        else:
                            return z + '/'
                    else:
                        return obecnasciezk + '/'
    else:
        return os.getcwd() + '/'

def open_file():
    """Open a file for editing."""
    filepath = askopenfilename(
        filetypes=[("Text Files", "*.txt")],
        initialdir=(cfgfile())
    )
    if not filepath:
        return
    txt_edit.delete("1.0", tk.END)
    with open(filepath, mode="r", encoding="utf-8") as input_file:
        text = input_file.read()
        txt_edit.insert(tk.END, text)
    odsw()


def save_file():
    """Save the current file as a new file."""
    filepath = asksaveasfilename(
        defaultextension=".txt",
        filetypes=[("Text Files", "*.txt")],
        initialdir=(cfgfile())
    )
    if not filepath:
        return
    with open(filepath, mode="w", encoding="utf-8") as output_file:
        text = txt_edit.get("1.0", tk.END)
        output_file.write(text)
    odsw()


def close_win():
    window.destroy()


def odsw():
    klista = listazfolderow()
    listbox.delete(0, tk.END)
    for x in klista:
        listbox.insert(tk.END, x)


def showcontent(event):
    x = listbox.curselection()
    file = listbox.get(x)
    with open(cfgfile() + file, 'r', encoding="utf-8") as file:
        file = file.read()
    txt_edit.delete('1.0', tk.END)
    txt_edit.insert(tk.END, file)
    zwroc_dane = file
    return zwroc_dane


def wybrany_plik():
    wybor_z_listy = listbox.curselection()
    nazwa_pliku_z_listy = listbox.get(wybor_z_listy)
    #nazwa_pliku_z_listy_bez =
    return nazwa_pliku_z_listy

def nazwapliku_bez_txt():
    nazwa_pliku = wybrany_plik()
    koncowka_nazwy = ".txt"
    nazwa_bez_txt = nazwa_pliku.removesuffix(koncowka_nazwy)
    '''if nazwa_pliku.endswith(".txt"):
        zmiennapodzielona = nazwa_pliku.rsplit('.')
        nazwa_bez_txt = zmiennapodzielona[0]'''
    return nazwa_bez_txt.replace('_', ' ')


def otworz_mail():
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    wiadomosc = showcontent('')
    mail.Subject = str(nazwapliku_bez_txt())
    mail.GetInspector
    index = mail.HTMLBody.find('>', mail.HTMLBody.find('<body'))
    mail.HTMLBody = mail.HTMLBody[:index + 1] + wiadomosc + mail.HTMLBody[index + 1:]
    mail.Display()


def export():
    otworz_mail()


def listazfolderow():
    if os.path.exists(cfgfile()):
        sciezka = os.listdir(cfgfile())
    else:
        sciezka = os.listdir()
    y = []
    for x in sciezka:
        if ".txt" in x:
            y.extend([x])
        if ".TXT" in x:
            y.extend([x])
    return y


def sprawdzsciezke():
    zmienna = cfgfile()
    zmienna2 = os.path.exists('cfg.ini')
    if not zmienna2:
        pass
    elif str(zmienna) == str(os.getcwd() + '/'):
        tkinter.messagebox.showwarning(
            title="Błąd wprowadzonej ścieżki",
            message=f"Ścieżka w pliku cfg.ini błędna. Ustawiono {zmienna}"
        )


def brakplikucfg():
    zmienna = os.path.exists('cfg.ini')
    if zmienna != True:
        zmienna2 = tkinter.messagebox.askyesno(
            title="Brak pliku cfg",
            message="Czy utworzyć plik cfg"
        )
        if zmienna2:
            with open('cfg.ini', 'w') as fp:
                pass


var1 = listazfolderow()
window = tk.Tk()
window.title("Mail template")
zmiennalisty = tk.Variable(value=var1)
window.rowconfigure(0)
window.columnconfigure(1, minsize=200, weight=1)
txt_edit = tk.Text(window)
frm_buttons = tk.Frame(window, bg='darkgray', background="black")
frm_list = tk.Frame(window, bd=1)
btn_open = tk.Button(frm_buttons, text="Open", command=open_file)
btn_save = tk.Button(frm_buttons, text="Save As...", command=save_file)
btn_close = tk.Button(frm_buttons, text="Close", command=close_win)
btn_refresh = tk.Button(frm_buttons, text="Refresh", command=odsw)
btn_export = tk.Button(frm_buttons, text="Export", command=export)

brakplikucfg()
sprawdzsciezke()

listbox = tk.Listbox(
    frm_list,
    listvariable=zmiennalisty,
    selectmode=tk.SINGLE,
    width=50

)



scrollbar = ttk.Scrollbar(
    frm_list,
    orient=tk.VERTICAL,
    command=listbox.yview
)

pokaz_liste = showcontent(event='')
listbox.bind("<<ListboxSelect>>", pokaz_liste)


frm_buttons.grid(row=0, column=0, sticky="n")
btn_open.grid(row=0, column=0, sticky="ew")
btn_save.grid(row=0, column=1, sticky="ew")
btn_close.grid(row=0, column=2, sticky="ew")
btn_refresh.grid(row=1, column=0, sticky="ew")
btn_export.grid(row=1, column=1, sticky="ew")
frm_list.grid(row=1, column=0, sticky="NSEW")
listbox.grid(row=0, column=0, sticky="NSEW")
scrollbar.grid(row=0, column=1, sticky="NS")
txt_edit.grid(row=1, column=1, sticky="NSEW")


window.mainloop()

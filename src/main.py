import tkinter as tk
from tkinter import *
from tkinter import ttk, messagebox
import csv
import openpyxl
from datetime import date
import re
import shutil


with open('Справочники/setup.csv', mode='r', encoding='utf-8') as file:
    reader = csv.reader(file, delimiter=";")
    org_data = next(reader)
org_data.insert(1, date.today().strftime("%d.%m.%Y"))

root = tk.Tk()
root.title("ЭСМ-7. Страница 1")
##root.geometry("1070x500+460+250")
root.geometry("1070x500")


def validate_integer(P):
    pattern = '0123456789'
    for char in P:
        if char not in pattern:
            return False
    return True


def validate_date(P):
    pattern = '0123456789.'
    for char in P:
        if char not in pattern:
            return False
    return True


def validate_code(P):
    pattern = '0123456789-'
    for char in P:
        if char not in pattern:
            return False
    return True


vcmd = (root.register(validate_integer), '%P')
vcmtdate = (root.register(validate_date), '%P')
vcmtcode = (root.register(validate_code), '%P')

labels = [
    "ОКУД",
    "Дата составления",
    "Организация",
    "Адрес организации",
    "Телеофон организации",
    "ОКПО"
]

doc_entries = []
tk.Label(root, text="Справка №").grid(row=0, column=0, padx=10, pady=5, sticky=NW)

entry = tk.Entry(root, validate="key", validatecommand=vcmd)
entry.grid(row=0, column=1, padx=10, pady=5, sticky=NW)
doc_entries.append(entry)

for i, label in enumerate(labels):
    tk.Label(root, text=label).grid(row=i + 1, column=0, padx=10, pady=5, sticky=NW)
    entry = tk.Entry(root)
    entry.insert(0, org_data[i])
    entry.config(state='disabled')
    entry.grid(row=i + 1, column=1, padx=10, pady=5, sticky=NW)
    doc_entries.append(entry)

doc_labels = [
    "Фирма-заказчик",
    "Юр. адрес заказчика",
    "Телеофон заказчика",
    "ОКПО заказчика"
]

for i, label in enumerate(doc_labels):
    # print(i)
    tk.Label(root, text=label).grid(row=i, column=2, padx=10, pady=5, sticky=NW)
    if i == 3:
        entry = tk.Entry(root, validate="key", validatecommand=vcmd)
    elif i == 2:
        entry = tk.Entry(root, validate="key", validatecommand=vcmd)
    else:
        entry = tk.Entry(root, validate="key")
    entry.grid(row=i, column=3, padx=10, pady=5, sticky=NW)
    doc_entries.append(entry)

doc_labels = [
    "Наименование объекта",
    "Адрес объекта"
]

for i, label in enumerate(doc_labels):
    tk.Label(root, text=label).grid(row=i + 5, column=2, padx=10, pady=5, sticky=NW)
    entry = tk.Entry(root, validate="key")
    entry.grid(row=i + 5, column=3, padx=10, pady=5, sticky=NW)
    doc_entries.append(entry)

doc_labels = [
    "Машина",
    "Машинисты",
]
for i, label in enumerate(doc_labels):
    tk.Label(root, text=label).grid(row=i + 8, column=2, padx=10, pady=5, sticky=NW)


with open('Справочники/mashines.csv', mode='r', encoding='utf-8') as file:
    reader = csv.reader(file, delimiter=";")
    mashines = ["\t".join(m) for m in reader]

    cm = ttk.Combobox(state="readonly", values=mashines)
    cm.grid(row=8, column=3, padx=10, pady=5, sticky=NW)
    doc_entries.append(cm)

with open('Справочники/stuff.csv', mode='r', encoding='utf-8') as file:
    reader = csv.reader(file, delimiter=";")
    stuff = [m[0] for m in reader]
    cm2 = ttk.Combobox(state="readonly", values=stuff)
    cm2.grid(row=9, column=3, padx=10, pady=5, sticky=NW)
    doc_entries.append(cm2)

doc_labels = [
    "Код вида операции",
    "Дата начала работ",
    "Дата окончания работ"
]

frame = ttk.Frame(borderwidth=1, relief=SOLID, padding=[8, 10])
frame.grid(row=7, column=0, padx=10, pady=5, sticky=NW, columnspan=2)

for i in range(3):
    ttk.Label(frame, text=doc_labels[i]).grid(row=0, column=i, padx=10, pady=5)
    if i == 0:
        entry = ttk.Entry(frame, validate="key", validatecommand=vcmtcode, justify=CENTER)
        entry.insert(0, "-")
    else:
        entry = tk.Entry(frame, validate="key", validatecommand=vcmtdate, justify=CENTER)
    entry.grid(row=1, column=i, padx=10, pady=5)
    doc_entries.append(entry)

dop_labels = [
        "Заказчик (должность)",
        "Исполнитель (должность)"
    ]

for i, label in enumerate(dop_labels):
    tk.Label(root, text=label).grid(row=i + 8, column=0, padx=10, pady=5, sticky=NW)
    entry = tk.Entry(root, validate="key")
    entry.grid(row=i + 8, column=1, padx=10, pady=5, sticky=NW)
    doc_entries.append(entry)


def open_second_window():
    dt = doc_entries[16].get()
    if not re.match(pattern="[0-9]{2}.[0-9]{2}.[0-9]{4}", string=dt):
        messagebox.showinfo(title="Информация", message=f"Неверный формат даты: {dt}")
        return

    dt1 = doc_entries[17].get()
    if not re.match(pattern="[0-9]{2}.[0-9]{2}.[0-9]{4}", string=dt1):
        messagebox.showinfo(title="Информация", message=f"Неверный формат даты: {dt1}")
        return

    if not all([i.get() for i in doc_entries]):
        messagebox.showinfo(title="Информация", message="Заполните все поля!")
        return

    root.withdraw()
    second_window = tk.Toplevel(root)
    second_window.title("ЭСМ-7. Страница 2")

    ##root.geometry("1070x500+460+250")
    root.geometry("1070x500")
    works = []

    tk.Label(second_window, text="Код и наименование работы").grid(row=0, column=0, padx=10, pady=5, sticky=NW)
    # tk.Label(second_window, text="Код вида работы").grid(row=1, column=0, padx=10, pady=5, sticky=NW)
    tk.Label(second_window, text="Отработано машино-часов").grid(row=1, column=0, padx=10, pady=5, sticky=NW)
    tk.Label(second_window, text="Стоимость одного часа работы").grid(row=2, column=0, padx=10, pady=5, sticky=NW)

    with open('Справочники/works.csv', mode='r', encoding='utf-8') as file:
        reader = csv.reader(file, delimiter=";")
        stuff = ["; ".join(m) for m in reader]
        # tk.Label(second_window, text="Причина простоя:").grid(row=2, column=2, padx=10, pady=5, sticky=NW)
        comwr = ttk.Combobox(second_window, state="readonly", values=stuff)
        comwr.grid(row=0, column=1, padx=10, pady=5)

    for j in range(1, 3):
        # entry = tk.Entry(second_window, validate="key", validatecommand=vcmd)
        entry = tk.Entry(second_window, validate="key", validatecommand=vcmd)
        entry.grid(row=j, column=1, padx=10, pady=5, sticky=NW)
        works.append(entry)

    columns = ("Вид работы", "Код", "Отработано, м-ч", "Руб/м-ч", "Всего, руб")

    tree = ttk.Treeview(second_window, columns=columns, show="headings")
    tree.grid(row=4, column=0, padx=10, pady=5, sticky=NW, columnspan=4)

    for col in columns:
        tree.heading(col, text=col)
        tree.column(col, width=170)

    tree.insert(parent="", index=END, iid="A", values=("Итого:", "", 0, "", 0))
    tree.insert(parent="", index=END, iid="B", values=("Простои:", "", 0, 0, 0))
    tree.insert(parent="", index=END, iid="C", values=("Всего:", "", 0, "", 0))
    tree.insert(parent="", index=END, iid="D", values=("", "", "", "Сумма НДС", 0))
    tree.insert(parent="", index=END, iid="E", values=("", "", "", "Всего с НДС", 0), tag="gray")

    tree.tag_configure('gray', font=("TkDefaultFont", 13, "bold"))
    with open('Справочники/downtime.csv', mode='r', encoding='utf-8') as file:
        reader = csv.reader(file, delimiter=";")
        stuff = [", ".join(m) for m in reader]
        tk.Label(second_window, text="Причина простоя:").grid(row=0, column=2, padx=10, pady=5, sticky=NW)
        compr = ttk.Combobox(second_window, state="readonly", values=stuff)
        compr.grid(row=0, column=3, padx=10, pady=5, sticky=NW)

    tk.Label(second_window, text="Время простоя по вине заказчика:").grid(row=1, column=2, padx=10, pady=5, sticky=NW)
    tk.Label(second_window, text="Стоимость часа простоя:").grid(row=2, column=2, padx=10, pady=5, sticky=NW)
    prost = []
    for j in range(1, 3):
        entry = tk.Entry(second_window, validate="key", validatecommand=vcmd)
        entry.grid(row=j, column=3, padx=10, pady=5, sticky=NW)
        prost.append(entry)

    def screen_table():
        val_A = list(tree.item("A").values())[2]

        times = int(works[0].get()) + int(val_A[2])
        prices = int(works[0].get()) * int(works[1].get()) + int(val_A[4])
        cvc, wrk = comwr.get().split("; ")

        tree.insert("", tk.END, values=tuple([wrk, cvc] + [i.get() for i in works] + [int(works[0].get()) * int(works[1].get())]))
        tree.item("A", values=("Итого:", "", times, "", prices))
        val_B = list(tree.item("B").values())[2]

        tree.item("C", values=("Всего:", "", times + int(val_B[2]), "", prices + int(val_B[4])))
        tree.item("D", values=("", "", "", "Сумма НДС", round((prices + int(val_B[4])) * 0.2, ndigits=2)))
        tree.item("E", values=("", "", "", "Всего с НДС", round((prices + int(val_B[4])) * 1.2, ndigits=2)))

        tree.move("A", parent="", index=END)
        tree.move("B", parent="", index=END)
        tree.move("C", parent="", index=END)
        tree.move("D", parent="", index=END)
        tree.move("E", parent="", index=END)

        for entry in works:
            entry.delete(0, END)
        comwr.set("")

    def prostoi():
        t, p = int(prost[0].get()), int(prost[1].get())
        tree.item("B", values=("Простои:", compr.get()[:2], t, p, t * p))
        val_C = list(tree.item("C").values())[2]

        tree.item("C", values=("Всего:", "", t + int(val_C[2]), "", t * p + int(val_C[4])))
        tree.item("D", values=("", "", "", "Сумма НДС", (t * p + int(val_C[4])) * 0.2))
        tree.item("E", values=("", "", "", "Всего с НДС", (t * p + int(val_C[4])) * 1.2))

    def final():
        if len(tree.get_children("")) <= 5:
            messagebox.showinfo(title="Информация", message=f"Введите хотя бы одну услугу")
            return

        filename = f"Справка №{doc_entries[0].get()}.xlsx"
        shutil.copy("template.xlsx", filename)
        wb = openpyxl.load_workbook(filename)
        sheet = wb.active

        sheet["AW5"].value = doc_entries[0].get() # справка №
        sheet["CK8"].value = doc_entries[1].get() # ОКУД
        sheet["CK9"].value = doc_entries[2].get() # дата составления
        sheet["N10"].value = f"{doc_entries[3].get()}, {doc_entries[4].get()}, {doc_entries[5].get()}" # организация, адрес, телефон
        sheet["CK10"].value = doc_entries[6].get() # ОКПО

        sheet["J12"].value = f"{doc_entries[7].get()}, {doc_entries[8].get()}, {doc_entries[9].get()}" # заказчик, адрес, телефон
        sheet["CK11"].value = doc_entries[10].get()
        sheet["I14"].value = f"{doc_entries[11].get()}, {doc_entries[12].get()}" # Наименование объекта, Адрес объекта

        car, brand, number = doc_entries[13].get().split("\t")
        sheet["J16"].value = car # Наименование машины
        sheet["BF16"].value = brand # Марка
        sheet["AH18"].value = number # Номерной знак машины
        sheet["M19"].value = doc_entries[14].get() # Машинисты

        sheet["BY16"].value = doc_entries[15].get() # Код вида операции
        sheet["CK16"].value = doc_entries[16].get() # Дата начала работ
        sheet["CT16"].value = doc_entries[17].get() # Дата окончания работ

        sheet["U41"].value = doc_entries[18].get() # Заказчик (должность)
        sheet["U43"].value = doc_entries[19].get() # Исполнитель (должность)
        # -=-=-=-=-=-=-=-=

        ch = tree.get_children("")
        for j in range(len(ch) - 5):
            t = tree.item(ch[j])["values"]

            sheet[f"A{26 + j}"].value = t[0]
            sheet[f"AR{26 + j}"].value = t[1]
            sheet[f"BD{26 + j}"].value = t[2]
            sheet[f"BU{26 + j}"].value = t[3]
            sheet[f"CL{26 + j}"].value = t[4]

        sheet[f"BD33"].value = list(tree.item("A").values())[2][2]
        sheet[f"CL33"].value = list(tree.item("A").values())[2][4]

        sheet[f"AR34"].value = str(list(tree.item("B").values())[2][1]).rjust(2, "0")
        sheet[f"BD34"].value = list(tree.item("B").values())[2][2]
        sheet[f"BU34"].value = list(tree.item("B").values())[2][3]
        sheet[f"CL34"].value = list(tree.item("B").values())[2][4]

        sheet[f"BD35"].value = list(tree.item("C").values())[2][2]
        sheet[f"CL35"].value = list(tree.item("C").values())[2][4]

        sheet[f"CL36"].value = list(tree.item("D").values())[2][4]
        sheet[f"CL37"].value = list(tree.item("E").values())[2][4]

        if int(list(tree.item("C").values())[2][2]) <= 400:
            with open('Справочники/numbers.csv', mode='r', encoding='utf-8') as file:
                reader = csv.reader(file, delimiter=";")
                for _ in range(int(list(tree.item("C").values())[2][2]) + 1):
                    numb = next(reader)[1]
            sheet["H39"].value = numb # Пропись
        else:
            sheet["H39"].value = int(list(tree.item("C").values())[2][2])

        wb.save(filename)
        wb.close()
        second_window.withdraw()

    tk.Button(second_window, text="Добавить", command=screen_table).grid(row=3, column=0, columnspan=2, pady=10)
    tk.Button(second_window, text="Добавить простои", command=prostoi).grid(row=3, column=2, columnspan=2, pady=10)

    tk.Button(second_window, text="Экспорт", command=final).grid(row=5, column=0, columnspan=4, pady=10)


tk.Button(root, text="Далее", command=open_second_window).grid(row=15, column=0, columnspan=4, pady=10)
root.mainloop()

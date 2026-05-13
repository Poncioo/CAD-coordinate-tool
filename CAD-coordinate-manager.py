import tkinter as tk
from tkinter import Text, filedialog, ttk
import pandas as pd
import win32com.client
import pythoncom
import numpy as np


# Inizializza CAD
cad = win32com.client.Dispatch("NanoCAD.Application")
doc = cad.ActiveDocument
df = None
CAD_COM = {
    "AutoCAD": "AutoCAD.Application",
    "NanoCAD": "NanoCAD.Application",
    "BricsCAD": "BricsCAD.Application",
    "ZWCAD": "ZWCAD.Application",
}

DECIMAL_PLACES = 4
opzioni_decimali = ["0.00", "0.000", "0.0000", "0.00000", "0.000000", ]


def selezione_cad(event=None):
    global cad, doc
    scelta = dropdown_button6.get()
    com_str = CAD_COM.get(scelta)
    if com_str:
        try:
            cad = win32com.client.Dispatch(com_str)
            doc = cad.ActiveDocument
            root.title(scelta)

        except Exception as e:
            root.title("Non connesso")


def aggiorna_decimali(event=None):
    global DECIMAL_PLACES
    DECIMAL_PLACES = dropdown_decimali.current() + 2  # indice 0 = ".2f"


def format_col(df, col):
    fmt = f"{{:.{DECIMAL_PLACES}f}}"
    df[col] = df[col].apply(lambda x: fmt.format(x))
    return df


# Esporta le coordinate degli oggetti selezionati
def run_export():
    global df, data
    cad.Visible = True
    selezione = doc.PickfirstSelectionSet
    export_data = "Nessuna coordinata trovata"
    data = []
    counter = 0

    def get_non_zero_z(z):
        return z if z != 0 else ""

    for obj in selezione:
        oen = obj.EntityName

        if oen == "AcDbCircle":
            x, y, z = obj.Center
            counter += 1
            data.append([counter, x, y, get_non_zero_z(z)])

        elif oen == "AcDbPoint":
            x, y, z = obj.Coordinates
            counter += 1
            data.append([counter, x, y, get_non_zero_z(z)])

        elif oen in ["AcDbBlock", "AcDbBlockReference"]:
            x, y, z = obj.InsertionPoint
            counter += 1
            data.append([counter, x, y, get_non_zero_z(z)])

        elif oen in ["AcDbText", "AcDbMText"]:
            x, y, z = obj.InsertionPoint
            data.append([obj.TextString, x, y, get_non_zero_z(z)])

        elif oen == "AcDbLine":
            counter += 1
            x1, y1, z1 = obj.StartPoint
            x2, y2, z2 = obj.EndPoint
            data.append([f"{counter}.1", x1, y1, get_non_zero_z(z1)])
            data.append([f"{counter}.2", x2, y2, get_non_zero_z(z2)])

        elif oen == "AcDbPolyline":
            punti = tuple(obj.Coordinates[x:x + 2] for x in range(0, len(obj.Coordinates), 2))
            is_closed = len(punti) >= 2 and punti[0] == punti[-1]
            for i, punto in enumerate(punti):
                x, y = punto
                counter += 1
                if is_closed and i == len(punti) - 1:
                    counter -= 1
                    continue
                data.append([counter, x, y, ""])

        elif oen == "AcDb3dPolyline":
            punti = tuple(obj.Coordinates[x:x + 3] for x in range(0, len(obj.Coordinates), 3))
            is_closed = len(punti) >= 2 and punti[0] == punti[-1]
            for i, punto in enumerate(punti):
                x, y, z = punto
                counter += 1
                if is_closed and i == len(punti) - 1:
                    counter -= 1
                    continue
                data.append([counter, x, y, get_non_zero_z(z)])

    df = pd.DataFrame(data, columns=None, index=None)

    if data:
        df[0] = df[0].astype(str)
        for col in range(1, len(df.columns)):
            if df[col].dtype == np.float64:
                format_col(df, col)
        export_data = "\n".join(["\t".join(map(lambda x: "" if x == "nan" else str(x), row)) for row in df.values])
        print(export_data)
    else:
        print("Nessuna coordinata trovata")

    output_text.delete(1.0, tk.END)
    output_text.insert(tk.END, export_data)


# Importa coordinate correnti nel CAD
def run_import():
    def point(x, y, z):
        return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (x, y, z))

    if df is not None:
        df[0] = df[0].astype(str)
        print(df)
        for index, row in df.iterrows():
            if row[3] == "":
                row[3] = 0
            doc.ModelSpace.AddCircle(Center=point(float(row[1]), float(row[2]), float(row[3])), Radius=0.1)
            doc.ModelSpace.AddText(TextString=row[0], InsertionPoint=point(row[1], row[2], row[3]), Height="0.4")
    else:
        print("Nessuna coordinata trovata")
        output_text.delete(1.0, tk.END)
        output_text.insert(tk.END, "Nessuna coordinata trovata")


# Copia coordinate negli appunti
def copy_to_clipboard():
    text_to_copy = output_text.get(1.0, tk.END)
    root.clipboard_clear()
    root.clipboard_append(text_to_copy)
    root.update()


# Salva coordinate su file
def salva_su_file():
    if df is not None:
        file_path = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Text files", "*.txt")])
        if not file_path:
            return
        df.to_csv(file_path, sep='\t', columns=None, header=None, index=None)
    else:
        print('nessun dato trovato.')
        output_text.delete(1.0, tk.END)
        output_text.insert(tk.END, "Nessuna coordinata trovata")


# Apri file di coordinate
def apri_file():
    global df
    file_path = filedialog.askopenfilename(defaultextension=".txt", filetypes=[("Text files", "*.txt")])
    if not file_path:
        return
    df = pd.read_csv(file_path, header=None, index_col=None, delimiter="\t")
    print(df)
    df[0] = df[0].astype(str)
    for col in range(1, len(df.columns)):
        if df[col].dtype == np.float64:
            format_col(df, col)
    file_data = "\n".join(["\t".join(map(lambda x: "" if x == "nan" else str(x), row)) for row in df.values])
    output_text.delete(1.0, tk.END)
    output_text.insert(tk.END, file_data)


# Creazione della finestra principale
root = tk.Tk()
root.title("CAD coordinates manager")

# Creazione dei pulsanti
button1 = tk.Button(root, text="Esporta coordinate", command=run_export)
button2 = tk.Button(root, text="Importa coordinate", command=run_import)
button3 = tk.Button(root, text="Copia contenuto", command=copy_to_clipboard)
button4 = tk.Button(root, text="Apri file", command=apri_file)
button5 = tk.Button(root, text="Salva file", command=salva_su_file)
dropdown_button6 = ttk.Combobox(root, values=list(CAD_COM.keys()), state="readonly")
dropdown_button6.bind("<<ComboboxSelected>>", selezione_cad)
dropdown_button6.set("NanoCAD")
dropdown_decimali = ttk.Combobox(root, values=opzioni_decimali, state="readonly")
dropdown_decimali.set("0.0000")
dropdown_decimali.bind("<<ComboboxSelected>>", aggiorna_decimali)



# Finestra per visualizzare le coordinate
def on_delete(event):
    if event.keysym in {"Delete", "BackSpace"}:
        return "break"


output_text = Text(root, wrap=tk.WORD, width=40, height=20)
output_text.config(state="normal")
output_text.bind("<Key>", on_delete)
output_text.pack()

# Posizionamento dei pulsanti
button1.pack()
button2.pack()
button3.pack(side=tk.LEFT)
button4.pack(side=tk.RIGHT)
button5.pack(side=tk.RIGHT)
dropdown_button6.pack(side=tk.RIGHT)
dropdown_decimali.pack(side=tk.RIGHT)

# Avvio della GUI
root.mainloop()

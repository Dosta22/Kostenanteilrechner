import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime
import os

# 1. Daten aus Excel laden
def lade_daten():
    dateiname = "Anteile Anlagen- & Leistungsspezifisch.xlsx"
    try:
        df = pd.read_excel(dateiname)
        return dict(zip(df.iloc[:, 0], df.iloc[:, 1]))
    except Exception as e:
        print(f"Fehler beim Laden der Quelldatei: {e}")
        return {}

daten_dict = lade_daten()

# 2. Logik & Export-Funktionen
def berechne():
    try:
        rechnungsbetrag = float(entry_betrag.get().replace(',', '.'))
        auswahl_name = combo_anlagen.get()
        
        if not auswahl_name:
            messagebox.showwarning("Hinweis", "Bitte zuerst eine Anlagenart wählen!")
            return
        
        anteilMCK1 = float(daten_dict.get(auswahl_name))
        anteilMCK4 = 100 - anteilMCK1
        
        # Fee Logik
        if rechnungsbetrag < 1000:
            fee_proz = 5.0
        elif rechnungsbetrag > 5000:
            fee_proz = 7.5
        else:
            fee_proz = 8.5
            
        fee = rechnungsbetrag / 100 * fee_proz
        gesamt = rechnungsbetrag + fee
        
        # Berechnungen Anteile
        res = {
            "g1": gesamt / 100 * anteilMCK1,
            "r1": rechnungsbetrag / 100 * anteilMCK1,
            "f1": fee / 100 * anteilMCK1,
            "g4": gesamt / 100 * anteilMCK4,
            "r4": rechnungsbetrag / 100 * anteilMCK4,
            "f4": fee / 100 * anteilMCK4
        }

        # Text für die Anzeige im Fenster
        ausgabe = (
            f"BERECHNUNG VOM {datetime.now().strftime('%d.%m.%Y %H:%M')}\n"
            f"Anlage: {auswahl_name}\n"
            f"Netto-Rechnung: {rechnungsbetrag:.2f}€\n"
            f"Management Fee ({fee_proz}%): {fee:.2f}€\n"
            f"Gesamtbetrag: {gesamt:.2f}€\n"
            f"{'='*40}\n"
            f"ANTEIL MCK1 ({anteilMCK1:.2f}%):\n"
            f"Gesamtanteil:      {res['g1']:.2f}€\n"
            f"Rechnungsanteil:   {res['r1']:.2f}€\n"
            f"Fee-Anteil:        {res['f1']:.2f}€\n"
            f"{'-'*40}\n"
            f"ANTEIL MCK4 ({anteilMCK4:.2f}%):\n"
            f"Gesamtanteil:      {res['g4']:.2f}€\n"
            f"Rechnungsanteil:   {res['r4']:.2f}€\n"
            f"Fee-Anteil:        {res['f4']:.2f}€"
        )
        label_ausgabe.config(text=ausgabe)
        btn_export_txt.config(state="normal")
        btn_export_excel.config(state="normal")

        # Daten für Excel-Log zwischenspeichern (global für die Export-Funktion)
        global aktuell_daten
        aktuell_daten = {
            "Datum": datetime.now().strftime('%d.%m.%Y %H:%M'),
            "Anlage": auswahl_name,
            "Netto_Rechnung": rechnungsbetrag,
            "Fee_Prozent": fee_proz,
            "Fee_Euro": fee,
            "Gesamt": gesamt,
            "MCK1_Anteil_Prozent": anteilMCK1,
            "MCK1_Gesamt": res['g1'],
            "MCK4_Anteil_Prozent": anteilMCK4,
            "MCK4_Gesamt": res['g4']
        }

    except ValueError:
        messagebox.showerror("Fehler", "Bitte einen gültigen Betrag eingeben!")

def export_txt():
    inhalt = label_ausgabe.cget("text")
    zeitstempel = datetime.now().strftime("%Y-%m-%d_%H-%M")
    dateiname = filedialog.asksaveasfilename(defaultextension=".txt", initialfile=f"Berechnung_{zeitstempel}.txt")
    if dateiname:
        with open(dateiname, "w", encoding="utf-8") as f:
            f.write(inhalt)
        messagebox.showinfo("Erfolg", "Textdatei gespeichert!")

def export_to_log():
    log_file = "Berechnungs_Log.xlsx"
    new_data = pd.DataFrame([aktuell_daten])
    
    try:
        if os.path.exists(log_file):
            # Bestehende Datei laden und neue Zeile anhängen
            with pd.ExcelWriter(log_file, engine="openpyxl", mode="a", if_sheet_exists="overlay") as writer:
                try:
                    existing_df = pd.read_excel(log_file)
                    updated_df = pd.concat([existing_df, new_data], ignore_index=True)
                    updated_df.to_excel(writer, index=False)
                except:
                    new_data.to_excel(writer, index=False)
        else:
            # Neue Datei erstellen
            new_data.to_excel(log_file, index=False)
        
        messagebox.showinfo("Erfolg", f"Daten wurden in {log_file} gespeichert!")
    except Exception as e:
        messagebox.showerror("Fehler", f"Konnte Log nicht schreiben (evtl. Datei geöffnet?): {e}")

# 3. GUI
root = tk.Tk()
root.title("Kostenrechner Pro MCK1/MCK4")
root.geometry("600x800")

tk.Label(root, text="Netto Rechnungsbetrag (€):", font=("Arial", 10, "bold")).pack(pady=10)
entry_betrag = tk.Entry(root, font=("Arial", 12), width=20)
entry_betrag.pack()

tk.Label(root, text="Anlagenart auswählen:", font=("Arial", 10, "bold")).pack(pady=10)
combo_anlagen = ttk.Combobox(root, values=list(daten_dict.keys()), state="readonly", width=50)
combo_anlagen.pack()

# Buttons
btn_frame = tk.Frame(root)
btn_frame.pack(pady=20)

tk.Button(btn_frame, text="1. Berechnen", command=berechne, bg="#d1e7dd", width=18).grid(row=0, column=0, padx=5)
btn_export_txt = tk.Button(btn_frame, text="2. Als TXT speichern", command=export_txt, state="disabled", width=18)
btn_export_txt.grid(row=0, column=1, padx=5)
btn_export_excel = tk.Button(btn_frame, text="3. In Excel-Log schreiben", command=export_to_log, state="disabled", width=18, bg="#e2e2e2")
btn_export_excel.grid(row=0, column=2, padx=5)

label_ausgabe = tk.Label(root, text="Werte eingeben...", justify="left", font=("Consolas", 10), 
                         bg="white", relief="sunken", width=65, height=25, anchor="nw", padx=15, pady=15)
label_ausgabe.pack(pady=10)

root.mainloop()

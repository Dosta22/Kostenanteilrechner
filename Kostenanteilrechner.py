import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox

# 1. Daten aus Excel laden
def lade_daten():
    dateiname = "Anteile Anlagen- & Leistungsspezifisch.xlsx"
    try:
        df = pd.read_excel(dateiname)
        return dict(zip(df.iloc[:, 0], df.iloc[:, 1]))
    except Exception as e:
        print(f"Fehler beim Laden: {e}")
        return {}

daten_dict = lade_daten()

# 2. Die Berechnungslogik mit ALLER Ausgaben
def berechne():
    try:
        rechnungsbetrag = float(entry_betrag.get().replace(',', '.'))
        auswahl_name = combo_anlagen.get()
        
        if not auswahl_name:
            messagebox.showwarning("Hinweis", "Bitte zuerst eine Anlagenart wählen!")
            return
        
        # Prozentsatz aus Excel holen
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
        
        # BERECHNUNGEN MCK1 (wie im Original)
        gesamtMCK1 = gesamt / 100 * anteilMCK1
        rechnungMCK1 = rechnungsbetrag / 100 * anteilMCK1
        feeMCK1 = fee / 100 * anteilMCK1
        
        # BERECHNUNGEN MCK4 (wie im Original)
        gesamtMCK4 = gesamt / 100 * anteilMCK4
        rechnungMCK4 = rechnungsbetrag / 100 * anteilMCK4
        feeMCK4 = fee / 100 * anteilMCK4

        # Zusammenbau des Textes für das Fenster
        ergebnis_text = (
            f"ERGEBNISSE:\n"
            f"Rechnungsbetrag: {rechnungsbetrag:.2f}€\n"
            f"Fee ({fee_proz:.2f}%): {fee:.2f}€\n"
            f"Gesamtbetrag: {gesamt:.2f}€\n"
            f"{'='*35}\n"
            f"MCK1: {anteilMCK1:.2f}%\n"
            f"Gesamtanteil MCK1: {gesamtMCK1:.2f}€\n"
            f"Rechnungsanteil MCK1: {rechnungMCK1:.2f}€\n"
            f"Fee-Anteil MCK1: {feeMCK1:.2f}€\n"
            f"{'-'*35}\n"
            f"MCK4: {anteilMCK4:.2f}%\n"
            f"Gesamtanteil MCK4: {gesamtMCK4:.2f}€\n"
            f"Rechnungsanteil MCK4: {rechnungMCK4:.2f}€\n"
            f"Fee-Anteil MCK4: {feeMCK4:.2f}€"
        )
        label_ausgabe.config(text=ergebnis_text)

    except ValueError:
        messagebox.showerror("Fehler", "Bitte einen gültigen Betrag eingeben!")

# 3. GUI SETUP
root = tk.Tk()
root.title("Kostenanteil Rechner MCK1 & MCK4")
root.geometry("500x650") # Fenster etwas größer für alle Daten

tk.Label(root, text="Netto Rechnungsbetrag (€):", font=("Arial", 10, "bold")).pack(pady=10)
entry_betrag = tk.Entry(root, font=("Arial", 12), width=20)
entry_betrag.pack()

tk.Label(root, text="Anlagenart auswählen:", font=("Arial", 10, "bold")).pack(pady=10)
combo_anlagen = ttk.Combobox(root, values=list(daten_dict.keys()), state="readonly", width=45)
combo_anlagen.pack()

tk.Button(root, text="Berechnen", command=berechne, bg="#d1e7dd", font=("Arial", 10, "bold"), height=2, width=15).pack(pady=20)

# Das Ausgabefeld mit weißem Hintergrund und Rahmen
label_ausgabe = tk.Label(root, text="Werte eingeben und Berechnen klicken", justify="left", 
                         font=("Consolas", 10), bg="white", relief="sunken", 
                         width=55, height=18, anchor="nw", padx=10, pady=10)
label_ausgabe.pack(pady=10)

root.mainloop()

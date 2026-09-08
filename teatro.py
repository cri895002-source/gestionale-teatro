import os
import sys
from datetime import datetime

try:
    import tkinter as tk
    from tkinter import messagebox, ttk, simpledialog
    import pandas as pd
except ImportError:
    import tkinter as tk
    from tkinter import messagebox
    root = tk.Tk()
    root.withdraw()
    messagebox.showerror("Errore", "Mancano librerie! Scrivi nel terminale:\npip install pandas")
    sys.exit()

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_FILE = os.path.join(BASE_DIR, "prenotazioni_teatro.csv")

# --- CONFIGURAZIONE PROFILI UTENTE ---
UTENTI_CONFIG = {
    "daniela": "daniela2026",
    "elisa": "elisa2026",
    "mirirena": "mirirena2026",
    "amministratore": "admin123"
}

UTENTE_CORRENTE = None

# Tariffe distinte per il 1° Settore (Platea) e il 2° Settore (Tribuna)
PREZZI = {
    "Platea_Intero": "18.00",
    "Platea_Ridotto": "12.00",
    "Tribuna_Intero": "15.00",
    "Tribuna_Ridotto": "10.00",
    "Speciale": "5.00"
}

COLONNE_DB = [
    "ID_Posto", "Settore", "Fila", "Posto", "Nome", "Cognome", 
    "Telefono", "Prezzo", "Ridotto", "Disabilita", 
    "Anticipo_Pagato", "Metodo_Pagamento", "Scadenza_Acconto", "Data_Ora_Prenotazione", "Operatore"
]

if not os.path.exists(DB_FILE) or os.path.getsize(DB_FILE) == 0:
    df_init = pd.DataFrame(columns=COLONNE_DB)
    df_init.to_csv(DB_FILE, index=False)

def esegui_login():
    global UTENTE_CORRENTE
    login_w = tk.Tk()
    login_w.title("LOGIN OPERATORE")
    login_w.geometry("350x220")
    login_w.configure(bg="#2c3e50")
    login_w.eval('tk::PlaceWindow . center')
    
    tk.Label(login_w, text="🔐 ACCESSO SISTEMA TEATRO", bg="#2c3e50", fg="white", font=("Arial", 11, "bold")).pack(pady=15)
    
    tk.Label(login_w, text="Nome Utente:", bg="#2c3e50", fg="white").pack()
    e_user = tk.Entry(login_w, font=("Arial", 10), width=25)
    e_user.pack(pady=2)
    
    tk.Label(login_w, text="Password:", bg="#2c3e50", fg="white").pack()
    e_pass = tk.Entry(login_w, font=("Arial", 10), width=25, show="*")
    e_pass.pack(pady=2)
    
    def verifica():
        global UTENTE_CORRENTE
        u = e_user.get().strip().lower()
        p = e_pass.get()
        
        if u in UTENTI_CONFIG and UTENTI_CONFIG[u] == p:
            UTENTE_CORRENTE = u
            login_w.destroy()
        else:
            messagebox.showerror("Errore", "Nome utente o Password errati!")
            
    tk.Button(login_w, text="ACCEDI", command=verifica, bg="#2ecc71", fg="white", font=("Arial", 10, "bold"), width=15).pack(pady=15)
    login_w.mainloop()

esegui_login()

if not UTENTE_CORRENTE:
    sys.exit()

root = tk.Tk()
root.title(f"SISTEMA GESTIONALE TEATRO - Operatore in uso: {UTENTE_CORRENTE.upper()}")
root.geometry("1400x980") 
root.configure(bg="#2c3e50")

pulsanti_posti = {}
lettere_file = [chr(i) for i in range(ord('A'), ord('S') + 1)]

def salva_csv(df):
    try:
        df.to_csv(DB_FILE, index=False)
        return True
    except PermissionError:
        messagebox.showerror("File Bloccato", "Chiudi Excel prima di salvare!")
        return False

def apri_excel():
    if os.path.exists(DB_FILE):
        try:
            os.startfile(DB_FILE)
        except Exception as e: 
            messagebox.showerror("Errore", f"Impossibile aprire Excel: {e}")
    else:
        messagebox.showinfo("Cassa", "Database vuoto.")

def reset_totale():
    if messagebox.askyesno("ATTENZIONE", "Vuoi svuotare tutto il teatro?"):
        df_init = pd.DataFrame(columns=COLONNE_DB)
        if salva_csv(df_init):
            for id_u, btn in pulsanti_posti.items():
                f = id_u.split('_')[0]
                btn.configure(bg='#1abc9c' if lettere_file.index(f) < 11 else '#3498db')
            messagebox.showinfo("Reset", "Teatro svuotato con successo!")

def mostra_cassa():
    try: 
        df = pd.read_csv(DB_FILE)
    except: 
        return
    if df.empty: 
        messagebox.showinfo("Cassa", "Nessuna prenotazione presente.")
        return
    df['Prezzo_Num'] = pd.to_numeric(df['Prezzo'].astype(str).str.replace(',', '.'), errors='coerce').fillna(0)
    tot_posti = len(df)
    tot_incasso = df['Prezzo_Num'].sum()
    report = f"📊 RESOCONTO CASSA\n\nPosti Occupati: {tot_posti} / 456\nIncasso Totale: {tot_incasso:.2f} €\n\nDettaglio Metodi:\n"
    for metodo, sub_df in df.groupby('Metodo_Pagamento'):
        sub_df['Prezzo_Num'] = pd.to_numeric(sub_df['Prezzo'].astype(str).str.replace(',', '.'), errors='coerce').fillna(0)
        report += f"• {metodo}: {sub_df['Prezzo_Num'].sum():.2f} €\n"
    messagebox.showinfo("Resoconto Cassa", report)

def apri_calcolatrice():
    calc_w = tk.Toplevel(root)
    calc_w.title("Calcolatrice")
    calc_w.geometry("300x400")
    calc_w.configure(bg="#2c3e50")
    calc_w.resizable(False, False)
    
    schermo = tk.Entry(calc_w, font=("Arial", 20), justify="right", bd=10, insertwidth=4, width=14, bg="#ecf0f1")
    schermo.grid(row=0, column=0, columnspan=4, padx=10, pady=10)
    
    def click_tasto(carattere):
        schermo.insert(tk.END, carattere)
        
    def cancella():
        schermo.delete(0, tk.END)
        
    def calcola():
        try:
            risultato = eval(schermo.get().replace('x', '*').replace(':', '/'))
            cancella()
            schermo.insert(0, f"{risultato:.2f}" if isinstance(risultato, float) else str(risultato))
        except:
            cancella()
            schermo.insert(0, "Errore")

    pulsanti = [
        ('7', 1, 0), ('8', 1, 1), ('9', 1, 2), (':', 1, 3),
        ('4', 2, 0), ('5', 2, 1), ('6', 2, 2), ('x', 2, 3),
        ('1', 3, 0), ('2', 3, 1), ('3', 3, 2), ('-', 3, 3),
        ('C', 4, 0), ('0', 4, 1), ('.', 4, 2), ('+', 4, 3)
    ]
    
    for (testo, r, c) in pulsanti:
        cmd = cancella if testo == 'C' else (lambda t=testo: click_tasto(t))
        col_b = "#e74c3c" if testo == 'C' else ("#3498db" if testo in [':','x','-','+'] else "#34495e")
        tk.Button(calc_w, text=testo, font=("Arial", 14, "bold"), bg=col_b, fg="white", width=4, height=2, command=cmd).grid(row=r, column=c, padx=5, pady=5)
        
    tk.Button(calc_w, text="=", font=("Arial", 14, "bold"), bg="#2ecc71", fg="white", width=22, height=2, command=calcola).grid(row=5, column=0, columnspan=4, padx=5, pady=5)

def cerca_spettatore():
    cerca_w = tk.Toplevel(root)
    cerca_w.title("Cerca Spettatore")
    cerca_w.geometry("480x250")
    cerca_w.configure(bg="#34495e")
    tk.Label(cerca_w, text="Inserisci Nome o Cognome:", bg="#34495e", fg="white", font=("Arial", 10, "bold")).pack(pady=10)
    entry_r = tk.Entry(cerca_w, font=("Arial", 11), width=30)
    entry_r.pack(pady=5)
    def esegui():
        t = entry_r.get().strip().lower()
        if not t: return
        try: 
            df = pd.read_csv(DB_FILE)
        except: 
            return
        if df.empty: return
        ris = df[df['Nome'].astype(str).str.lower().str.contains(t) | df['Cognome'].astype(str).str.lower().str.contains(t)]
        if ris.empty: 
            messagebox.showinfo("Ricerca", "Nessun risultato trovato.")
        else:
            msg = "🔍 RISULTATI RICERCA:\n\n"
            for _, r in ris.iterrows():
                op_stampa = str(r.get('Operatore', 'N/D')).upper()
                prezzo_p = str(r.get('Prezzo', '0.00'))
                msg += f"👤 {r['Nome']} {r['Cognome']} ➔ [{r['Settore']}] Fila {r['Fila']} Posto {r['Posto']} | Prezzo: {prezzo_p} € | Op: [{op_stampa}]\n"
            messagebox.showinfo("Esito Ricerca", msg)
            cerca_w.destroy()
    tk.Button(cerca_w, text="🔍 CERCA", command=esegui, bg="#2980b9", fg="white", font=("Arial", 10, "bold")).pack(pady=10)

def apri_legenda():
    leg_w = tk.Toplevel(root)
    leg_w.title("Legenda")
    leg_w.geometry("380x280")
    leg_w.configure(bg="#34495e")
    tk.Label(leg_w, text="📋 INFO SETTORI E COLORI TEATRO", bg="#34495e", fg="white", font=("Arial", 11, "bold")).pack(pady=10)
    txt = "🟢 Verde = 1° Settore (Platea)\n🔵 Azzurro = 2° Settore (Tribuna)\n🟡 Giallo = Prenotato senza acconto\n🔴 Rosso = Prenotato e PAGATO\n👑 Fila A/B = Riservate\n♿ Fila L = Disabili"
    tk.Label(leg_w, text=txt, bg="#34495e", fg="white", justify="left").pack(padx=20, pady=5)
    tk.Button(leg_w, text="CHIUDI", command=leg_w.destroy).pack(pady=10)

def aggiorna_testo_banner_prezzi():
    lbl_info_costi.config(
        text=f"💰 TARIFFE BASE  ➔  Platea (1° Settore): {PREZZI['Platea_Intero']}€ / {PREZZI['Platea_Ridotto']}€  |  "
             f"Tribuna (2° Settore): {PREZZI['Tribuna_Intero']}€ / {PREZZI['Tribuna_Ridotto']}€"
    )

def modifica_prezzi_interfaccia():
    if UTENTE_CORRENTE != "amministratore":
        messagebox.showerror("Accesso Negato", "Solo l'AMMINISTRATORE può modificare il listino prezzi!")
        return
        
    p_int = simpledialog.askstring("Platea (1° Settore)", f"Prezzo Intero PLATEA (Attuale: {PREZZI['Platea_Intero']} €):")
    p_rid = simpledialog.askstring("Platea (1° Settore)", f"Prezzo Ridotto PLATEA (Attuale: {PREZZI['Platea_Ridotto']} €):")
    t_int = simpledialog.askstring("Tribuna (2° Settore)", f"Prezzo Intero TRIBUNA (Attuale: {PREZZI['Tribuna_Intero']} €):")
    t_rid = simpledialog.askstring("Tribuna (2° Settore)", f"Prezzo Ridotto TRIBUNA (Attuale: {PREZZI['Tribuna_Ridotto']} €):")
    
    if p_int: PREZZI["Platea_Intero"] = p_int.replace("€", "").strip()
    if p_rid: PREZZI["Platea_Ridotto"] = p_rid.replace("€", "").strip()
    if t_int: PREZZI["Tribuna_Intero"] = t_int.replace("€", "").strip()
    if t_rid: PREZZI["Tribuna_Ridotto"] = t_rid.replace("€", "").strip()
    
    aggiorna_testo_banner_prezzi()
    messagebox.showinfo("Successo", "Listino prezzi aggiornato con successo!")

def apri_scheda(id_unico, settore, fila, posto, lista_posti, dati_esistenti, modifica_attiva):
    win = tk.Toplevel(root)
    win.title(f"Scheda Prenotazione - [{settore}] Fila {fila}")
    win.geometry("470x780")
    win.configure(bg="#f8f9fa")

    op_registrazione = dati_esistenti.get("Operatore", UTENTE_CORRENTE).upper()
    tk.Label(win, text=f"📍 SETTORE: {settore.upper()} | OPERATORE: {op_registrazione}", bg="#34495e", fg="white", font=("Arial", 10, "bold"), pady=6).pack(fill="x")

    tk.Label(win, text="Nome:", bg="#f8f9fa").pack(anchor="w", padx=15, pady=2)
    e_nome = tk.Entry(win, width=38, font=("Arial", 10)); e_nome.insert(0, dati_esistenti.get("Nome","")); e_nome.pack(padx=15, pady=2)
    
    tk.Label(win, text="Cognome:", bg="#f8f9fa").pack(anchor="w", padx=15, pady=2)
    e_cognome = tk.Entry(win, width=38, font=("Arial", 10)); e_cognome.insert(0, dati_esistenti.get("Cognome","")); e_cognome.pack(padx=15, pady=2)
    
    tk.Label(win, text="Telefono:", bg="#f8f9fa").pack(anchor="w", padx=15, pady=2)
    e_tel = tk.Entry(win, width=38, font=("Arial", 10)); e_tel.insert(0, dati_esistenti.get("Telefono","")); e_tel.pack(padx=15, pady=2)
    
    prezzo_base_default = PREZZI["Platea_Intero"] if settore == "Platea" else PREZZI["Tribuna_Intero"]
    prezzo_iniziale = dati_esistenti.get("Prezzo", prezzo_base_default)

    tk.Label(win, text="✏️ PREZZO MODIFICABILE PER SINGOLO POSTO (€):", bg="#f8f9fa", fg="#c0392b", font=("Arial", 9, "bold")).pack(anchor="w", padx=15, pady=(8, 2))
    e_prezzo = tk.Entry(win, width=15, font=("Arial", 11, "bold"), bg="#fffde7"); e_prezzo.insert(0, prezzo_iniziale); e_prezzo.pack(anchor="w", padx=15, pady=2)
    
    frame_rapido = tk.Frame(win, bg="#f8f9fa")
    frame_rapido.pack(anchor="w", padx=15, pady=2)
    
    def imposta_prezzo(val, rid_val):
        e_prezzo.delete(0, tk.END)
        e_prezzo.insert(0, val)
        c_rid.set(rid_val)
        aggiorna_riepilogo_costi()

    p_int_set = PREZZI["Platea_Intero"] if settore == "Platea" else PREZZI["Tribuna_Intero"]
    p_rid_set = PREZZI["Platea_Ridotto"] if settore == "Platea" else PREZZI["Tribuna_Ridotto"]

    tk.Button(frame_rapido, text=f"Carica Base Intero ({p_int_set}€)", font=("Arial", 8), command=lambda: imposta_prezzo(p_int_set, "NO")).pack(side="left", padx=2)
    tk.Button(frame_rapido, text=f"Carica Base Ridotto ({p_rid_set}€)", font=("Arial", 8), command=lambda: imposta_prezzo(p_rid_set, "SÌ")).pack(side="left", padx=2)
    tk.Button(frame_rapido, text=f"Speciale ({PREZZI['Speciale']}€)", font=("Arial", 8), command=lambda: imposta_prezzo(PREZZI["Speciale"], "SÌ")).pack(side="left", padx=2)

    frame_costi = tk.LabelFrame(win, text="💶 Riepilogo Calcolo Costo", bg="#ebedef", font=("Arial", 9, "bold"), padx=10, pady=5)
    frame_costi.pack(fill="x", padx=15, pady=8)

    lbl_dettaglio_costi = tk.Label(frame_costi, text="", bg="#ebedef", justify="left", font=("Consolas", 9, "bold"), fg="#2c3e50")
    lbl_dettaglio_costi.pack(anchor="w")

    def aggiorna_riepilogo_costi(*args):
        try:
            p_val = float(e_prezzo.get().replace(",", "."))
        except ValueError:
            p_val = 0.0
        
        testo_dettaglio = f"Settore: {settore}\n"
        totale = 0.0
        for p in lista_posti:
            testo_dettaglio += f"• Fila {fila} - Posto {p}: {p_val:.2f} €\n"
            totale += p_val
            
        testo_dettaglio += f"-----------------------------\nTOTALE ({len(lista_posti)} posti): {totale:.2f} €"
        lbl_dettaglio_costi.config(text=testo_dettaglio)

    e_prezzo.bind("<KeyRelease>", aggiorna_riepilogo_costi)

    tk.Label(win, text="Ridotto?", bg="#f8f9fa").pack(anchor="w", padx=15, pady=2)
    c_rid = ttk.Combobox(win, values=["NO", "SÌ"], state="readonly", width=10)
    c_rid.set(dati_esistenti.get("Ridotto","NO"))
    c_rid.pack(padx=15, pady=2)

    tk.Label(win, text="Disabilità?", bg="#f8f9fa").pack(anchor="w", padx=15, pady=2)
    c_dis = ttk.Combobox(win, values=["NO", "SÌ"], state="readonly", width=10)
    c_dis.set(dati_esistenti.get("Disabilita", "SÌ" if fila in ['A', 'L'] else "NO"))
    c_dis.pack(padx=15, pady=2)

    tk.Label(win, text="Anticipo Pagato?", bg="#f8f9fa").pack(anchor="w", padx=15, pady=2)
    c_ant = ttk.Combobox(win, values=["NO", "SÌ"], state="readonly", width=10)
    c_ant.set(dati_esistenti.get("Anticipo_Pagato","NO"))
    c_ant.pack(padx=15, pady=2)

    tk.Label(win, text="Metodo Pagamento:", bg="#f8f9fa").pack(anchor="w", padx=15, pady=2)
    c_met = ttk.Combobox(win, values=["Contanti", "Carta", "Bancomat"], state="readonly", width=20)
    c_met.set(dati_esistenti.get("Metodo_Pagamento","Contanti"))
    c_met.pack(padx=15, pady=2)

    tk.Label(win, text="Scadenza Acconto (es. 20/06):", bg="#f8f9fa").pack(anchor="w", padx=15, pady=2)
    e_scad = tk.Entry(win, width=15); e_scad.insert(0, dati_esistenti.get("Scadenza_Acconto","")); e_scad.pack(padx=15, pady=2)
 
    data_ora_corrente = dati_esistenti.get("Data_Ora_Prenotazione", "")
    op_esistente = dati_esistenti.get("Operatore", "")
    
    if data_ora_corrente:
        info_txt = f"📆 Assegnato il: {data_ora_corrente}"
        if op_esistente:
            info_txt += f" | Operatore: [{op_esistente.upper()}]"
        tk.Label(win, text=info_txt, fg="#7f8c8d", bg="#f8f9fa", font=("Arial", 9, "italic")).pack(pady=5)

    def salva():
        if not e_nome.get().strip() or not e_cognome.get().strip():
            messagebox.showerror("Errore", "Nome e Cognome obbligatori!")
            return
        try: 
            df_l = pd.read_csv(DB_FILE)
        except: 
            df_l = pd.DataFrame(columns=COLONNE_DB)
        
        ora_click = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        
        nuove_righe = []
        for p_num in lista_posti:
            id_c = f"{fila}_{p_num}"
            if id_c in pulsanti_posti:
                if "ID_Posto" in df_l.columns and not df_l.empty:
                    df_l = df_l[df_l["ID_Posto"] != id_c]
                
                data_salvataggio = data_ora_corrente if (modifica_attiva and data_ora_corrente) else ora_click
                operatore_salvataggio = op_esistente if (modifica_attiva and op_esistente) else UTENTE_CORRENTE
                
                nuove_righe.append({
                    "ID_Posto": id_c, "Settore": settore, "Fila": fila, "Posto": p_num, 
                    "Nome": e_nome.get().strip(), "Cognome": e_cognome.get().strip(), "Telefono": e_tel.get().strip(), 
                    "Prezzo": e_prezzo.get().strip(), "Ridotto": c_rid.get(), "Disabilita": c_dis.get(), 
                    "Anticipo_Pagato": c_ant.get(), "Metodo_Pagamento": c_met.get(), "Scadenza_Acconto": e_scad.get().strip(),
                    "Data_Ora_Prenotazione": data_salvataggio, "Operatore": operatore_salvataggio
                })
        if nuove_righe:
            if salva_csv(pd.concat([df_l, pd.DataFrame(nuove_righe)], ignore_index=True)):
                for p_num in lista_posti:
                    id_c = f"{fila}_{p_num}"
                    if id_c in pulsanti_posti:
                        pulsanti_posti[id_c].configure(bg='#e74c3c' if c_ant.get()=='SÌ' else '#f1c40f')
                win.destroy()
 
    tk.Button(win, text="💾 SALVA PRENOTAZIONE", command=salva, bg="#2ecc71", fg="white", font=("Arial", 10, "bold")).pack(pady=12)
    aggiorna_riepilogo_costi()
 
def gestisci_posto(id_unico, settore, fila, posto):
    try: 
        df = pd.read_csv(DB_FILE)
    except: 
        df = pd.DataFrame(columns=COLONNE_DB)
    
    if fila == 'A':
        if not messagebox.askyesno("⚠️ FILA A RISERVATA", "Stai modificando la FILA A (Autorità/Disabili).\nVuoi davvero procedere?"):
            return
 
    gia_prenotato = df[df["ID_Posto"] == id_unico] if ("ID_Posto" in df.columns and not df.empty) else pd.DataFrame()
    
    if gia_prenotato.empty:
        tipo = messagebox.askyesnocancel(f"Posto {fila}-{posto}", "Vuoi fare una prenotazione MULTIPLA?\n\n[SÌ] Più posti insieme\n[NO] Solo questo posto")
        if tipo is True:
            multi_w = tk.Toplevel(root)
            multi_w.title("Prenotazione Posti Multipli")
            multi_w.geometry("420x220")
            multi_w.configure(bg="#2c3e50")
            
            tk.Label(multi_w, text="⚠️ SEPARAZIONE CON VIRGOLA", bg="#2c3e50", fg="#f1c40f", font=("Arial", 11, "bold")).pack(pady=(10, 2))
            tk.Label(multi_w, text="Inserisci i numeri dei posti separati da VIRGOLA!\nEsempio: 13, 14, 15", bg="#2c3e50", fg="white", font=("Arial", 9, "bold")).pack(pady=5)
            
            tk.Label(multi_w, text=f"Settore: {settore} | Fila {fila}:", bg="#2c3e50", fg="white").pack(pady=2)
            entry_m = tk.Entry(multi_w, font=("Arial", 11), width=30)
            entry_m.insert(0, str(posto))
            entry_m.pack(pady=5)
            
            def ok():
                try:
                    lista = [int(p.strip()) for p in entry_m.get().split(",") if p.strip().isdigit()]
                    if not lista:
                        messagebox.showerror("Errore", "Inserisci numeri validi separati da virgola!")
                        return
                    multi_w.destroy()
                    apri_scheda(id_unico, settore, fila, posto, lista, {}, False)
                except Exception: 
                    messagebox.showerror("Errore Format", "Formato non valido! Usa la VIRGOLA (es: 13,14,15)")
                    
            tk.Button(multi_w, text="✓ CONFERMA SELEZIONE", command=ok, bg="#2ecc71", fg="white", font=("Arial", 10, "bold")).pack(pady=10)
            
        elif tipo is False:
            apri_scheda(id_unico, settore, fila, posto, [posto], {}, False)
    else:
        p = gia_prenotato.iloc[0]
        op_effettuato = str(p.get('Operatore', 'N/D')).upper()
        data_effettuata = str(p.get('Data_Ora_Prenotazione', 'N/D'))
        prezzo_p = str(p.get('Prezzo', '0.00'))
        
        info_msg = (
            f"📌 POSTO OCCUPATO: [{settore}] Fila {fila} - Posto {posto}\n"
            f"👤 Intestatario: {p['Nome']} {p['Cognome']}\n"
            f"💶 Costo Posto: {prezzo_p} €\n"
            f"🛠️ Operazione effettuata da: [{op_effettuato}]\n"
            f"📅 Data/Ora: {data_effettuata}\n\n"
            f"[SÌ] Modifica Scheda | [NO] Libera Posto"
        )
        
        scelta = messagebox.askyesnocancel(f"Gestione Posto {fila}-{posto}", info_msg)
        if scelta is True:
            apri_scheda(id_unico, settore, fila, posto, [posto], p.to_dict(), True)
        elif scelta is False:
            df = df[df["ID_Posto"] != id_unico]
            if salva_csv(df):
                pulsanti_posti[id_unico].configure(bg='#1abc9c' if settore=="Platea" else '#3498db')
 
# --- PANNELLO COMANDI SUPERIORE ---
top_panel = tk.Frame(root, bg="#34495e")
top_panel.pack(fill="x", pady=5)
tk.Label(top_panel, text="🎭 PALCOSCENICO", bg="#34495e", fg="white", font=("Arial", 14, "bold")).pack(side="left", padx=20, pady=10)

lbl_info_costi = tk.Label(top_panel, bg="#34495e", fg="#ecf0f1", font=("Arial", 9, "italic"))
lbl_info_costi.pack(side="left", padx=15)
aggiorna_testo_banner_prezzi()

tk.Button(top_panel, text="❌ AZZERA", command=reset_totale, bg="#c0392b", fg="white").pack(side="right", padx=5)
tk.Button(top_panel, text="⚙️ MODIFICA COSTI", command=modifica_prezzi_interfaccia, bg="#7f8c8d", fg="white").pack(side="right", padx=5)
tk.Button(top_panel, text="🧮 CALCOLATRICE", command=apri_calcolatrice, bg="#9b59b6", fg="white").pack(side="right", padx=5)
tk.Button(top_panel, text="📊 CASSA", command=mostra_cassa, bg="#27ae60", fg="white").pack(side="right", padx=5)
tk.Button(top_panel, text="🔍 CERCA", command=cerca_spettatore, bg="#2980b9", fg="white").pack(side="right", padx=5)
tk.Button(top_panel, text="🟢 EXCEL", command=apri_excel, bg="#27ae60", fg="white").pack(side="right", padx=5)
tk.Button(top_panel, text="📋 LEGENDA", command=apri_legenda, bg="#e67e22", fg="white").pack(side="right", padx=5)

# --- MAPPA SCROLLABILE ---
canvas = tk.Canvas(root, bg="#2c3e50", highlightthickness=0)
v_scroll = tk.Scrollbar(root, orient="vertical", command=canvas.yview)
h_scroll = tk.Scrollbar(root, orient="horizontal", command=canvas.xview)
map_frame = tk.Frame(canvas, bg="#2c3e50")
map_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
canvas.create_window((0, 0), window=map_frame, anchor="nw")
canvas.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)

canvas.pack(side="top", fill="both", expand=True, padx=10, pady=5)
v_scroll.pack(side="right", fill="y")
h_scroll.pack(side="bottom", fill="x")
 
try: 
    df_attuale = pd.read_csv(DB_FILE)
except: 
    df_attuale = pd.DataFrame(columns=COLONNE_DB)
 
riga_vis = 1
for r_idx, letter in enumerate(lettere_file):
    settore = "Platea" if r_idx < 11 else "Tribuna"
    colore_base = '#1abc9c' if settore == "Platea" else '#3498db'
    
    if letter == 'A':
        tk.Label(map_frame, text="🚪 EMERGENZA (SX)", bg="#c0392b", fg="white").grid(row=riga_vis, column=0)
        tk.Label(map_frame, text="EMERGENZA (DX) 🚪", bg="#c0392b", fg="white").grid(row=riga_vis, column=26)
        riga_vis += 1
    if letter == 'L': 
        tk.Label(map_frame, text="—" * 80, bg="#2c3e50", fg="#7f8c8d").grid(row=riga_vis, column=1, columnspan=24, pady=10)
        riga_vis += 1
        
    pref = "👑 " if letter in ['A','B'] else ("♿ " if letter == 'L' else "")
    tk.Label(map_frame, text=f"{pref}Fila {letter}", bg="#2c3e50", fg="white", font=("Arial", 9, "bold")).grid(row=riga_vis, column=0, padx=5)
    
    for c_idx in range(1, 24):
        id_u = f"{letter}_{c_idx}"
        col_g = c_idx if c_idx <= 12 else c_idx + 1
        if c_idx == 13:
            tk.Label(map_frame, text=" ", bg="#2c3e50").grid(row=riga_vis, column=13)
            
        row_db = df_attuale[df_attuale["ID_Posto"] == id_u] if ("ID_Posto" in df_attuale.columns and not df_attuale.empty) else pd.DataFrame()
        colore_f = colore_base if row_db.empty else ('#e74c3c' if row_db.iloc[0]['Anticipo_Pagato'] == 'SÌ' else '#f1c40f')
        
        btn = tk.Button(map_frame, text=str(c_idx), bg=colore_f, fg="white", font=("Arial", 8, "bold"), width=3,
                        command=lambda id_u=id_u, s=settore, f=letter, p=c_idx: gestisci_posto(id_u, s, f, p))
        btn.grid(row=riga_vis, column=col_g, padx=1, pady=1)
        pulsanti_posti[id_u] = btn
    
    tk.Label(map_frame, text=f"Fila {letter} {pref}", bg="#2c3e50", fg="white", font=("Arial", 9, "bold")).grid(row=riga_vis, column=26, padx=5)
    riga_vis += 1
 
tk.Label(map_frame, text="🚪 EMERGENZA (SX)", bg="#c0392b", fg="white").grid(row=riga_vis, column=0, pady=5)
tk.Label(map_frame, text="EMERGENZA (DX) 🚪", bg="#c0392b", fg="white").grid(row=riga_vis, column=26, pady=5)

# --- BARRA DI STATO IN BASSO CON OPERATORE E OROLOGIO (SENZA LOGO) ---
status_bar = tk.Frame(root, bd=1, relief="sunken", bg="#34495e")
status_bar.pack(side="bottom", fill="x")

lbl_operatore = tk.Label(status_bar, text=f"👤 Operatore in uso: {UTENTE_CORRENTE.upper()}", bg="#34495e", fg="#ecf0f1", font=("Arial", 9, "bold"))
lbl_operatore.pack(side="left", padx=10, pady=5)

lbl_orologio = tk.Label(status_bar, bg="#34495e", fg="#ecf0f1", font=("Arial", 9, "bold"))
lbl_orologio.pack(side="right", padx=10, pady=5)

def aggiorna_data_ora():
    stringa_data_ora = datetime.now().strftime("📅 %d/%m/%Y   🕒 %H:%M:%S")
    lbl_orologio.config(text=stringa_data_ora)
    lbl_orologio.after(1000, aggiorna_data_ora)

aggiorna_data_ora()

root.mainloop()
import openpyxl # per creare e modificare file Excel
import locale  # per impostare la lingua italiana nei nomi dei mesi
from datetime import datetime,timedelta 
from openpyxl.styles import Font, Alignment 

locale.setlocale(locale.LC_TIME, 'it_IT.UTF-8') # imposta locale italiano per nomi mesi (se disponibile)

def trova_domenica_di_pasqua(year: int) -> datetime: # calcolo la data della Pasqua (Gregorian Easter) per un dato anno 'anno'usando l'Algoritmo Anonimo (Gregoriano) (di Meeus/Jones/Butcher) (variante dell'algoritmo di Gauss). Algoritmo di Gauss, determina l'esatta data del giorno di Pasqua (che cade sempre la prima domenica dopo la prima luna piena successiva all'equinozio di primavera). Funzione prevista per restiruire un oggetto datetime (-> datetime SUGGERIMENTO DI TIPO FACOLTATIVO è una promessa che fai come sviluppatore sulla natura del valore restituito, rendendo il tuo codice più chiaro e facile da mantenere.)
    """
    Calcola la data della Pasqua (Gregoriana) usando l'algoritmo di Meeus/Jones/Butcher.
    Valido per gli anni successivi al 1583.
    """
    a = year % 19  # "Golden number" – posizione dell'anno nel ciclo metonico (approssima fasi lunari)
    b = year // 100  # Secolo.
    c = year % 100 #  Anno nel secolo.
    d = b // 4 # Correzioni per il ciclo solare (anni bisestili ogni 4 anni, con eccezioni secolari)
    e = b % 4  # ^     "           "
    # 3) correzioni secolari (f-g):
    f = (b + 8) // 25 # Correzione per l'equinozio di primavera (che si sposta lentamente nei secoli)
    g = (b - f + 1) // 3 # Ulteriore ajuste per la regola gregoriana (anni bisestili non ogni 100 anni, ma ogni 400)
    # 4) Epacta (h):
    h = (19 * a + b - d - g + 15) % 30 # Età della luna al 21 marzo (giorni dopo l'ultima luna nuova). È il cuore dell'algoritmo: approssima la prima luna piena dopo l'equinozio (fissato al 21 marzo).
    # 5) correzioni per l'anno nel secolo (i-k):
    i = c // 4 # Simili a d/e, ma per l'anno specifico.
    k = c % 4  # ^     "            "
    # 6) giorno della settimana (l):
    l = (32 + 2 * e + 2 * i - h - k) % 7 # Calcola quanti giorni dopo la luna piena cade la domenica (Pasqua è la prima domenica post-luna piena).
    # 7) correzione speciale (m):
    m = (a + 11 * h + 22 * l) // 451

    mese_finale = (h + l - 7 * m + 114) // 31 #  3 (marzo) o 4 (aprile)
    giorno_finale = ((h + l - 7 * m + 114) % 31) + 1 # giorno del mese
    return datetime(year, mese_finale, giorno_finale)


def crea_calendario_excel():
    today = datetime.now()
    this_month = today.month
    this_year = today.year
    month_name = datetime(this_year, this_month, 1).strftime('%B').capitalize() # nome del mese in italiano. String format time: converte un oggetto data leggibile dalla macchina in una stringa di testo formattata (%B = nome completo del mese secondo la lingua locale)
    pasqua_day = trova_domenica_di_pasqua(this_year)

    # Festività nazionali italiane
    feste_nazionali = [
        datetime(this_year, 1, 1),    # Capodanno
        datetime(this_year, 1, 6),    # Epifania
        pasqua_day,                   # Pasqua, da mia funzione
        pasqua_day + timedelta(days=1), # Pasquetta
        datetime(this_year, 4, 25),   # Liberazione
        datetime(this_year, 5, 1),    # Festa del Lavoro
        datetime(this_year, 6, 2),    # Festa della Repubblica
        datetime(this_year, 8, 15),   # Ferragosto
        datetime(this_year, 11, 1),   # Ognissanti
        datetime(this_year, 12, 8),   # Immacolata
        datetime(this_year, 12, 25),  # Natale
        datetime(this_year, 12, 26),  # Santo Stefano
    ]

    # Calcolo dell'ultimo giorno del mese corrente
    if this_month == 12: # gestisco il passaggio da Dicembre a Gennaio senza librerire aggiuntive
        next_month = 1
        next_year = this_year + 1
    else:
        next_month = this_month + 1
        next_year = this_year

    first_of_next_month = datetime(next_year, next_month, 1)
    last_of_this_month = first_of_next_month - timedelta(days=1)
    numb_month_days = last_of_this_month.day

    # Creazione file Excel
    workbook = openpyxl.Workbook() # creo nuovo file di lavoro (Workbook) Excel in memoria
    active_wb_reference = workbook.active # ottengo il riferimento al foglio di lavoro attivo
    active_wb_reference.title = month_name

    # Intestazione
    header_text = f"{month_name} {this_year}" # intestazione
    header_cell = active_wb_reference.cell(row=1, column=1, value=header_text)
    header_cell.font = Font(bold=True, size=18)

    # Stili
    feriale_black_style = Font(color="000000") # Default nero
    festivo_red_style = Font(color="FFFF0000")
    centered_alignment = Alignment(horizontal="center", vertical="center")
    week_initials = ("L", "M", "M", "G", "V", "S", "D")
    current_line = 4

    active_wb_reference.column_dimensions['A'].width = 5 

    # Popolamento dei giorni
    for day in range(1, numb_month_days + 1):
        current_date = datetime(this_year, this_month, day)
        day_of_week = current_date.weekday()
        day_initial = week_initials[day_of_week]
        day_number_and_initial_str = f"{day}  {day_initial}"
        is_weekend = day_of_week >= 5
        cell_style = festivo_red_style if (is_weekend or current_date in feste_nazionali) else feriale_black_style        

        day_number_and_initial_cell = active_wb_reference.cell(row=current_line, column=1, value=day_number_and_initial_str) # scrittura e formattazione della cella (colonna A)
        day_number_and_initial_cell.font = cell_style # è necessario assegnare allineamento e stile separatamente in openpyxl, perché sono proprietà diverse delle celle di Excel
        day_number_and_initial_cell.alignment = centered_alignment

        current_line += 1
    
    file_name = f"Calendario {month_name}_{this_year}.xlsx"
    # Devo salvare il file con metodo .save() e non posso creare il file con with open perchè il contenuto non è gestito come una sequenza semplice di caratteri o righe (stream), ma Un archivio compresso (simile a un file .zip) contenente molteplici file XML (per fogli, stili, formattazione, ecc.)
    try:
        workbook.save(file_name)
        print(f"✅ File Excel creato con successo: {file_name}")
    except PermissionError as e:
        print(f"❌ Errore durante il salvataggio del file: {e}")
    except Exception as e:
        print(f"❌ Errore durante il salvataggio del file: {e}")

# Se importo il mio script in un altro file, il codice non verrà eseguito automaticamente:
if __name__ == "__main__": # quando eseguo l'altro file.py, la condizione 'if __name__ == "__main__"' in calendario.py è FALSA. Il file Excel non viene creato. La funzione crea_calendario_excel è ora disponibile in altro_script.py come calendario.crea_calendario_excel().
    crea_calendario_excel() # Garantisce che le operazioni di modificano dello stato del sistema (come la creazione di un file o la stampa di un messaggio di benvenuto) avvengano solo quando l'utente intende eseguire quello specifico file.
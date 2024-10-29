import openpyxl


def carica_soluzioni(file_soluzioni):
    # Legge il file di testo con le soluzioni e i punti
    soluzioni = {}
    with open(file_soluzioni, 'r') as file:
        linee = file.readlines()

    # Organizza i dati in un dizionario {cella: (formula, punti)}
    for i in range(0, len(linee), 3):
        cella = linee[i].strip()  # La cella
        formula = linee[i + 1].strip()  # La formula
        punti = int(linee[i + 2].strip())  # I punti
        soluzioni[cella] = (formula, punti)

    return soluzioni


def controlla_formule(nome_file_excel, soluzioni):
    # Apre il file Excel e controlla le formule nel primo foglio
    workbook = openpyxl.load_workbook(nome_file_excel, data_only=False)
    foglio = workbook.active  # Usa il primo foglio
    punteggio_totale = 0
    risultati = {nome_file_excel: {}}

    for cella, (formula_attesa, punti) in soluzioni.items():
        valore_cella = foglio[cella].value

        # Verifica se la formula è corretta
        if valore_cella == formula_attesa:
            risultati[nome_file_excel][cella] = f"Formula corretta: {formula_attesa} (+{punti} punti)"
            punteggio_totale += punti
        else:
            risultati[nome_file_excel][
                cella] = f"Formula errata. Attesa: {formula_attesa}, Trovata: {valore_cella} (0 punti)"

    return risultati, punteggio_totale


def calcola_punteggio_totale(file_elenco_excel, file_soluzioni):
    # Carica le soluzioni e inizializza il punteggio complessivo
    soluzioni = carica_soluzioni(file_soluzioni)
    punteggio_totale_globale = 0
    risultati_globale = {}
    punteggi_per_file = {}

    # Legge la lista dei file Excel
    with open(file_elenco_excel, 'r') as file:
        nomi_file_excel = [line.strip() for line in file if line.strip()]

    # Calcola il punteggio per ciascun file Excel
    for nome_file_excel in nomi_file_excel:
        risultati, punteggio_totale = controlla_formule(nome_file_excel, soluzioni)
        risultati_globale.update(risultati)
        punteggi_per_file[nome_file_excel] = punteggio_totale

    return risultati_globale, punteggi_per_file


# Esempio di utilizzo
file_elenco_excel = 'elencoalunni.txt'
file_soluzioni = 'soluzioni.txt'
risultati_globale, punteggi_per_file = calcola_punteggio_totale(file_elenco_excel,
                                                                                          file_soluzioni)

# Stampa i risultati e il punteggio per ogni file
for nome_file, celle in risultati_globale.items():
    print(f"File: {nome_file}")
    for cella, risultato in celle.items():
        print(f"  Cella {cella}: {risultato}")
    print(f"Punteggio totale per {nome_file}: {punteggi_per_file[nome_file]}")
#print(f"\nPunteggio totale complessivo ottenuto: {punteggio_totale_globale}")

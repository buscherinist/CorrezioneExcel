import openpyxl

def carica_soluzioni(file_soluzioni):
    # Legge il file di testo con le soluzioni e i punti, organizzati per foglio
    soluzioni = {}
    with open(file_soluzioni, 'r') as file:
        linee = file.readlines()

    foglio_corrente = None
    i = 0
    while i < len(linee):
        line = linee[i].strip()

        # Se la linea è un nome di foglio
        if line.startswith("Foglio"):
            foglio_corrente = line
            soluzioni[foglio_corrente] = {}
            i += 1
            continue

        # Altrimenti, leggiamo cella, formula e punti
        if foglio_corrente:
            cella = line
            formula = linee[i + 1].strip()
            punti = int(linee[i + 2].strip())
            soluzioni[foglio_corrente][cella] = (formula, punti)
            i += 3

    return soluzioni


def controlla_formule(nome_file_excel, soluzioni):
    # Apre il file Excel e controlla le formule
    workbook = openpyxl.load_workbook(nome_file_excel, data_only=False)
    punteggio_totale = 0
    risultati = {nome_file_excel: {}}

    for foglio_nome, celle in soluzioni.items():
        foglio = workbook[foglio_nome]
        risultati[nome_file_excel][foglio_nome] = {}

        for cella, (formula_attesa, punti) in celle.items():
            valore_cella = foglio[cella].value

            # Verifica se la formula è corretta
            if valore_cella == formula_attesa:
                risultati[nome_file_excel][foglio_nome][cella] = f"Formula corretta: {formula_attesa} (+{punti} punti)"
                punteggio_totale += punti
            else:
                risultati[nome_file_excel][foglio_nome][
                    cella] = f"Formula errata. Attesa: {formula_attesa}, Trovata: {valore_cella} (0 punti)"

    return risultati, punteggio_totale

def calcola_punteggio_totale(file_elenco_excel, file_soluzioni):
    # Carica le soluzioni e inizializza il punteggio complessivo
    soluzioni = carica_soluzioni(file_soluzioni)
    risultati_globale = {}

    # Legge la lista dei file Excel
    with open(file_elenco_excel, 'r') as file:
        nomi_file_excel = [line.strip() for line in file if line.strip()]

    # Calcola il punteggio per ciascun file Excel
    for nome_file_excel in nomi_file_excel:
        risultati, punteggio_totale = controlla_formule("./verifiche/"+nome_file_excel, soluzioni)
        risultati_globale.update(risultati)

    return risultati_globale, punteggio_totale

# Main
file_elenco_excel = 'elencoalunni.txt'
file_soluzioni = 'soluzioni.txt'
risultati_globale, punteggio_totale = calcola_punteggio_totale(file_elenco_excel, file_soluzioni)

# Stampa i risultati e il punteggio totale complessivo
for nome_file, fogli in risultati_globale.items():
    print(f"File: {nome_file}")
    for foglio, celle in fogli.items():
        print(f"  Foglio: {foglio}")
        for cella, risultato in celle.items():
            print(f"    Cella {cella}: {risultato}")
    print(f"\nPunteggio totale complessivo ottenuto: {punteggio_totale}")

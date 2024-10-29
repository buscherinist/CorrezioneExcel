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
        if not line.startswith("=") and not line.isdigit() and line:
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
    print(soluzioni)
    return soluzioni


def controlla_formule(nome_file_excel, file_soluzioni):
    # Carica le soluzioni e il file Excel
    soluzioni = carica_soluzioni(file_soluzioni)
    workbook = openpyxl.load_workbook(nome_file_excel, data_only=False)

    # Controlla ogni cella e calcola il punteggio totale
    risultati = {}
    punteggio_totale = 0
    for foglio_nome, celle in soluzioni.items():
        foglio = workbook[foglio_nome]
        risultati[foglio_nome] = {}

        for cella, (formula_attesa, punti) in celle.items():
            valore_cella = foglio[cella].value

            # Verifica se la formula è corretta
            if valore_cella == formula_attesa:
                risultati[foglio_nome][cella] = f"Formula corretta: {formula_attesa} (+{punti} punti)"
                punteggio_totale += punti
            else:
                risultati[foglio_nome][
                    cella] = f"Formula errata. Attesa: {formula_attesa}, Trovata: {valore_cella} (0 punti)"

    return risultati, punteggio_totale


# Esempio di utilizzo
nome_file_excel = 'alunno.xlsx'
file_soluzioni = 'soluzioni2.txt'
risultati, punteggio_totale = controlla_formule(nome_file_excel, file_soluzioni)

# Stampa i risultati e il punteggio totale
for foglio, celle in risultati.items():
    print(f"Foglio: {foglio}")
    for cella, risultato in celle.items():
        print(f"  Cella {cella}: {risultato}")
print(f"\nPunteggio totale ottenuto: {punteggio_totale}")


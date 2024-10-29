import openpyxl
from sympy import sympify, simplify

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


def formule_uguali(formula1, formula2):
    try:
        # Converte le formule in espressioni simboliche
        expr1 = sympify(formula1)  # Sostituisci con il nome corretto
        expr2 = sympify(formula2)
        print("1")
        print(expr1, expr2)
        # Semplifica le espressioni
        simplified_expr1 = simplify(expr1)
        simplified_expr2 = simplify(expr2)
        print("2")
        print(simplified_expr1, simplified_expr2)

        # Confronta le espressioni semplificate
        return simplified_expr1 == simplified_expr2
    except Exception as e:
        print(f"Errore nel processamento delle formule: {e}")
        return False

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
            valore_cella = valore_cella.replace(" ", "")
            # Verifica se la formula è corretta
            uguali = formule_uguali(valore_cella, formula_attesa)
            #if valore_cella == formula_attesa:
            if uguali:
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
with open("correzione.txt", "w") as file:
    file.write(f"Correzione\n")
# Stampa i risultati e il punteggio totale complessivo
for nome_file, fogli in risultati_globale.items():
    with open("correzione.txt", "a") as file:
        file.write(f"File: {nome_file}\n")
        print(f"File: {nome_file}")
        for foglio, celle in fogli.items():
            file.write(f"  Foglio: {foglio}\n")
            print(f"  Foglio: {foglio}")
            for cella, risultato in celle.items():
                file.write(f"    Cella {cella}: {risultato}\n")
                print(f"    Cella {cella}: {risultato}")
        file.write(f"\nPunteggio totale complessivo ottenuto: {punteggio_totale}\n\n\n")
        print(f"\nPunteggio totale complessivo ottenuto: {punteggio_totale}")
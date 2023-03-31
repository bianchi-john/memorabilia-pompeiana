# Importing the modules
import csv
import openpyxl
import os


def datiAnagrafica(fileIn, fileOut):
    try:
        # Definisci gli header per il tuo file CSV
        headers = ['ID', 'Reperto', 'Materiale', 'Altezza', 'Lunghezza', 'Larghezza', 'Spessore', 'Diametro', 'Stato di Conservazione', 'Descrizione', 'Note',
                'Soggetto Iconografico', 'Cronologia', 'Nazione', 'Città', 'Status', 'Museo / Collezionista', '# inv.', 'Data di Acquisizione', 'Modalità di Acquisizione']

        # Apri il file CSV in modalità scrittura
        directory = "dataJohn"  # Specifico la directory
        filename = fileOut  # Specifico il nome del file

        if not os.path.exists(directory):  # Creo la directory se non esiste
            os.makedirs(directory)

        filePathAndName = os.path.join(
            directory, filename)  # Creo il full name path

        with open(filePathAndName, 'w', newline='') as file:
            # Creo un oggetto writer per scrivere nel file CSV
            writer = csv.writer(file)
            # Apro il file Excel
            workbook = openpyxl.load_workbook(fileIn)

            # Seleziono il foglio di lavoro che vuoi leggere
            worksheet = workbook['Dati Anagrafici']

            # Scrivo gli header nel file CSV
            writer.writerow(headers)

            # Leggo i dati nell'excel ciclicamente e li salvo in un oggetto per poi metterli nel CSV
            for row in worksheet.iter_rows(min_row=3, min_col=1, max_col=21):

                # Creo una lista vuota per contenere il contenuto delle celle
                cell_content = []

                # Itero su ogni cella e aggiungi il contenuto alla lista
                for cell in row:
                    cell_content.append(cell.value)

                # Stampo il contenuto
                # print (cell_content)

                # Se c'è qualche valora nella riga allora:
                if (cell_content[0]):
                    # Sostituisco il contenuto Null con la stringa 'Missing Value'
                    for i in range(len(cell_content)):
                        if cell_content[i] is None:
                            cell_content[i] = 'Missing Value'

                    # Scrivo il contenuto
                    writer.writerow([cell_content[0], cell_content[2], cell_content[3], cell_content[4], cell_content[5], cell_content[6], cell_content[7], cell_content[8], cell_content[9], cell_content[10],
                                    cell_content[11], cell_content[12], cell_content[13], cell_content[14], cell_content[15], cell_content[16], cell_content[17], cell_content[18], cell_content[19], cell_content[20]])
    except: 
        print('C\'è stato un errore duranta la creazione del file: \"' + fileOut + '\"')




def datiScavo(fileIn, fileOut):
    try:
        # Definisci gli header per il tuo file CSV
        headers = ['ID','ID SCAVO','Toponimo','Regio','Insula','Civico','Stanza','Informazioni di Dettaglio','Indicazioni Generali','Giorno','Mese','Anno','Soprintendente','Architetto Direttore','Note','Archivio','Segnatura','Riferimento PAH','Citazione']

        # Apri il file CSV in modalità scrittura
        directory = "dataJohn"  # Specifico la directory
        filename = fileOut  # Specifico il nome del file

        if not os.path.exists(directory):  # Creo la directory se non esiste
            os.makedirs(directory)

        filePathAndName = os.path.join(
            directory, filename)  # Creo il full name path

        with open(filePathAndName, 'w', newline='') as file:
            # Creo un oggetto writer per scrivere nel file CSV
            writer = csv.writer(file)
            # Apro il file Excel
            workbook = openpyxl.load_workbook(fileIn)

            # Seleziono il foglio di lavoro che vuoi leggere
            worksheet = workbook['Dati Anagrafici']

            # Scrivo gli header nel file CSV
            writer.writerow(headers)

            # Leggo i dati nell'excel ciclicamente e li salvo in un oggetto per poi metterli nel CSV
            for row in worksheet.iter_rows(min_row=3, min_col=1, max_col=19):

                # Creo una lista vuota per contenere il contenuto delle celle
                cell_content = []

                # Itero su ogni cella e aggiungi il contenuto alla lista
                for cell in row:
                    cell_content.append(cell.value)

                # Stampo il contenuto
                # print (cell_content)

                # Se c'è qualche valora nella riga allora:
                if (cell_content[0]):
                    # Sostituisco il contenuto Null con la stringa 'Missing Value'
                    for i in range(len(cell_content)):
                        if cell_content[i] is None:
                            cell_content[i] = 'Missing Value'

                    # Scrivo il contenuto
                    writer.writerow([cell_content[0], cell_content[2], cell_content[3], cell_content[4], cell_content[5], cell_content[6], cell_content[7], cell_content[8], cell_content[9], cell_content[10],
                                cell_content[11], cell_content[12], cell_content[13], cell_content[14], cell_content[15], cell_content[16], cell_content[17], cell_content[18]])
    except: 
        print('C\'è stato un errore duranta la creazione del file: \"' + fileOut + '\"')



datiScavo('excel.xlsx', 'dati_anagrafica_clean.csv')
datiAnagrafica('excel.xlsx', 'dati_scavo_clean.csv')

print ('DONE')

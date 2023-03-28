#Importing the modules
import openpyxl
from openpyxl import load_workbook
from openpyxl_image_loader import SheetImageLoader


def save_images(file_path, sheet_name):
    # Carica il file Excel e seleziona il foglio di lavoro
    wb = load_workbook(filename=file_path)
    sheet = wb[sheet_name]
    count = 0
    # Itera tutte le righe nel foglio di lavoro e salto le prime due perché non ci sono immagini
    for row in sheet.iter_rows():
        try:  
            count = count + 1
            if (count < 3):
                continue 
            my_string = str(row[1])
            subrow = my_string.split("'.",1)[1]
            subrow = subrow[:-1]
            print(subrow)
            #Carico il file excel e la sheet nella libreria per esportare le immagini
            pxl_doc = openpyxl.load_workbook(file_path)
            sheet = pxl_doc[sheet_name]
            #invoco l'image_loader
            image_loader = SheetImageLoader(sheet)
            #Prendo l'immagine in una specifica cella
            image = image_loader.get(subrow)
            rgb_im = image.convert('RGB')
            # La salvo in una directory
            rgb_im.save('my_path/' + subrow + '.jpg')
        except:
            print("Non è stata trovata un'immagine nella linea: " + str(row[1]))


save_images('excel.xlsx', 'Dati Anagrafici')

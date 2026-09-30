from db import get_connection
import csv
import re

cnxn = get_connection()
cursor = cnxn.cursor()
cursor2 = cnxn.cursor()

cursor.execute('select * from tbl_13029_STAGE_Articoli' )

# Rimuove tutti i caratteri superflui
def remove_bad_chars(string):
    string = string.replace(","," ")
    string = string.replace(":"," ")
    string = string.replace(";"," ")
    string = string.replace("\'"," ")
    
    # Rimuove spazi multipli
    string = re.sub(r'\s+', ' ', string)
    return string

# Creazione struttura dati per il lexicon. Usiamo un dizionario con key=parola e value=(score della parola)
lexicon = dict()

# Lettura lexicon
with open('lexicon.csv', 'r') as csvfile:
    reader = csv.reader(csvfile, delimiter=',')
    for row in reader:
        lexicon[row[0]] = int(row[1])

for row in cursor.fetchall():
    score = 0
    for word in remove_bad_chars(str(row.Titolo).lower()).split():
        if word in lexicon:
            score = score + lexicon[word]

    if (score > 0):
        cursor2.execute('UPDATE tbl_13029_STAGE_Articoli SET Etichetta = ? WHERE Id= ?', ('GRAVE', row.Id))
        print("articolo=%s, etichetta=%s, id=%s" % (row.Titolo, row.Etichetta, row.Id))
        cnxn.commit()
    else:
        cursor2.execute('UPDATE tbl_13029_STAGE_Articoli SET Etichetta = ? WHERE Id= ?', ('NON GRAVE', row.Id))
        print("articolo=%s, etichetta=%s, id=%s" % (row.Titolo, row.Etichetta, row.Id))
        cnxn.commit()
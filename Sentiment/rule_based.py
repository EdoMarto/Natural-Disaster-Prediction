from db import get_connection
import csv
import re

cnxn = get_connection()
cursor = cnxn.cursor()

cursor.execute('select * from tbl_13029_STAGE_Articoli ORDER BY DataArticolo DESC')

# Rimuove tutti i caratteri superflui
def remove_bad_chars(string):
    string = string.replace(","," ")
    string = string.replace(":"," ")
    string = string.replace(";"," ")
    string = string.replace("\'"," ")
    
    # Rimuove spazi multipli
    string = re.sub(r'\s+', ' ', string)
    return string

# CARICAMENTO ARTICOLI
articles = []

for row in cursor.fetchall():
    article = dict()
    article['titolo'] = remove_bad_chars(str(row.Titolo).lower())
    article['sottotitolo'] = remove_bad_chars(str(row.Sottotitolo).lower())
    article['descrizione'] = remove_bad_chars(str(row.Descrizione).lower())
    article['comune'] = str(row.Comune).lower()
    article['data'] = str(row.DataArticolo)
    article['tipo'] = str(row.TipoEvento).lower()
    article['etichetta'] = str(row.Etichetta)
    #print("article=%s, sottotitolo=%s, comune=%s" % (article['titolo'], article['sottotitolo'], article['comune']))
    articles.append(article)

# Creazione struttura dati per il lexicon. Usiamo un dizionario con key=parola e value=(score della parola)
lexicon = dict()

# Lettura lexicon
with open('lexicon.csv', 'r') as csvfile:
    reader = csv.reader(csvfile, delimiter=',')
    for row in reader:
        lexicon[row[0]] = int(row[1])

# Usare lexicon per classificare articoli
azzeccati = 0
for art in articles:
    score = 0
    for word in art['titolo'].split():
        if word in lexicon:
            score = score + lexicon[word]

    art['score'] = score
    if (score > 0):
        art['sentiment'] = 'GRAVE'
    else:
        art['sentiment'] = 'NON GRAVE'
        
    if(art['etichetta'] == art['sentiment']):
        azzeccati = azzeccati + 1

# Statistiche sugli articoli
total = float(len(articles))
num_pos = sum([1 for t in articles if t['sentiment'] == 'GRAVE'])
num_neg = sum([1 for t in articles if t['sentiment'] == 'NON GRAVE'])
print("GRAVI: %5d (%.1f%%)" % (num_pos, 100.0 * (num_pos/total)))
print("NON GRAVI: %5d (%.1f%%)" % (num_neg, 100.0 * (num_neg/total)))
print("ETICHETTE AZZECCATE: %5d (%.1f%%)" % (azzeccati, 100.0 * (azzeccati/total)))


# Stampa di alcuni articoli
articles_sorted = sorted(articles, key=lambda k: k['score'])

print("\n\nTOP ARTICOLI NON GRAVI")
negative_articles = [d for d in articles_sorted if d['sentiment'] == 'NON GRAVE']
for art in negative_articles[0:10]:
    print("score=%.2f, article=%s, tipo=%s" % (art['score'], art['titolo'], art['tipo']))

print("\n\nTOP ARTICOLI GRAVI")
positive_articles = [d for d in articles_sorted if d['sentiment'] == 'GRAVE']
for art in positive_articles[-10:]:
    print("score=%.2f, article=%s, tipo=%s" % (art['score'], art['titolo'], art['tipo']))

#print("\n\nTUTTI GLI ARTICOLI")
#for art in positive_articles:
#    print("score=%.2f, article=%s, tipo=%s, comune=%s" % (art['score'], art['titolo'], art['tipo'], art['comune']))
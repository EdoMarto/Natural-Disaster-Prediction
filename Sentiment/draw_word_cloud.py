import re
from matplotlib import pyplot as plt
import wordcloud
from db import get_connection

cnxn = get_connection()
cursor = cnxn.cursor()

cursor.execute('select * from tbl_13029_STAGE_Articoli ORDER BY DataArticolo DESC')

# CARICAMENTO ARTICOLI
articles = dict()
titles = []

# Rimuove tutti i caratteri superflui
def preprocess_text(string):
    string = string.replace(","," ")
    string = string.replace(":"," ")
    string = string.replace(";"," ")
    string = string.replace("\'"," ")
    
    # Rimuove spazi multipli
    string = re.sub(r'\s+', ' ', string)
    
    # Rimuove stop words
    stop_words = ["di","a","da","in","con","su","per","tra","fra","il","lo","la","i","gli","le","e","sul","sulla","sullo","dal","dalla","del","della","nel","nella","al","alla"]
    final = ""
    for word in string.split():
        if word not in stop_words:
            final = final + word + " "
    return final

for row in cursor.fetchall():
    titles.append(preprocess_text(row.Titolo.lower()))

# STAMPA WORDCLOUD
common_words=''
total_words=[]
for art in titles:
    common_words += art + " "
    if art not in total_words:
        total_words.append(art)
wordcloud = wordcloud.WordCloud().generate(common_words)
wordcloud.to_file('wordcloud.png')
plt.imshow(wordcloud, interpolation='bilinear')
plt.axis("off")
plt.show()
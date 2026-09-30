import pandas as pd
import numpy as np
import re
import nltk
from nltk.corpus import stopwords
from matplotlib import pyplot as plt

from numpy import array
from keras.preprocessing.text import one_hot
from keras.preprocessing.sequence import pad_sequences
from keras.models import Sequential
from keras.layers.core import Activation, Dropout, Dense
from keras.layers import Flatten
from keras.layers import GlobalMaxPooling1D
from tensorflow.keras.layers import LSTM,Dense, Dropout, SpatialDropout1D
from keras.layers.embeddings import Embedding
from sklearn.model_selection import train_test_split
from keras.preprocessing.text import Tokenizer
from keras import layers
import wordcloud
from sklearn.metrics import accuracy_score

from db import get_connection
import csv

import random

cnxn = get_connection()
cursor = cnxn.cursor()

cursor.execute('select * from tbl_13029_STAGE_Articoli ORDER BY DataArticolo DESC')

# CARICAMENTO ARTICOLI
articles = dict()
titles = []
sentiments = []

for row in cursor.fetchall():
    titles.append(row.Titolo.lower())
    sentiments.append(row.Etichetta)

articles['titolo'] = titles
articles['sentiment'] = sentiments
    
articles_df = pd.DataFrame.from_dict(articles)

articles_df.isnull().values.any()

articles_df.shape

print(articles_df.head())

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
    
X = []
sentences = list(articles_df['titolo'])
for sen in sentences:
    X.append(preprocess_text(sen))

y = articles_df['sentiment']

y = np.array(list(map(lambda x: 1 if x=="GRAVE" else 0, y)))

X_train = X
y_train = y

tokenizer = Tokenizer()
tokenizer.fit_on_texts(X_train)

X_train = tokenizer.texts_to_sequences(X_train)

max_words = 20

X_train = pad_sequences(X_train, padding='post', maxlen=max_words)

#-----------------------------------------------------------------
vocabulary_size = len(tokenizer.word_index) + 1

embedding_vector_length = 32
model = Sequential() 
model.add(Embedding(vocabulary_size, embedding_vector_length, input_length=max_words) )
model.add(SpatialDropout1D(0.25))
model.add(LSTM(50, dropout=0.5, recurrent_dropout=0.5))
model.add(Dropout(0.2))
model.add(Dense(1, activation='sigmoid')) 
model.compile(loss='binary_crossentropy',optimizer='adam', metrics=['accuracy'])  
print(model.summary()) 

batch_size = 32
num_epochs = 30

# selezionare numero epoche ottimale
from keras import callbacks
earlystopping = callbacks.EarlyStopping(monitor ="val_loss", 
                                        mode ="min", patience = 2, 
                                        restore_best_weights = True)
                                        
history = model.fit(X_train, y_train, validation_split=0.2, batch_size=batch_size, epochs=num_epochs, callbacks =[earlystopping])

# Esempio di articoli su cui fare previsione
sentence = ["bomba d acqua e grandine ancora allagamenti in sud sardegna", 
            "in settimana primo caldo estivo grazie all anticiclone africano",
            "perù violenta scossa di terremoto di magnitudo 7.2 nel sud del Paese"]
# conversione da testo a sequenza di int
sequences = tokenizer.texts_to_sequences(sentence)
# esecuzione padding sulla sequenza di interi
padded = pad_sequences(sequences, maxlen=max_words)
# Ottieni i label basandoti sulla probabilità -> if p>= 0.5 1 else 0
prediction = model.predict(padded)
pred_labels = []
for i in prediction:
    print(i)
    if i >= 0.5:
        pred_labels.append(1)
    else:
        pred_labels.append(0)
for i in range(len(sentence)):
    print(sentence[i])
    if pred_labels[i] == 1:
        s = 'GRAVE'
    else:
        s = 'NON GRAVE'
    print("Previsione : ",s)
    
# Grafico sull accuratezza
plt.plot(history.history['accuracy'])
plt.plot(history.history['val_accuracy'])
plt.title('model accuracy')
plt.ylabel('accuracy')
plt.xlabel('epoch')
plt.legend(['train', 'test'], loc='upper left')
plt.show()

# Grafico sulla loss
plt.plot(history.history['loss'])
plt.plot(history.history['val_loss'])
plt.title('model loss')
plt.ylabel('loss')
plt.xlabel('epoch')
plt.legend(['train', 'test'], loc='upper left')
plt.show()
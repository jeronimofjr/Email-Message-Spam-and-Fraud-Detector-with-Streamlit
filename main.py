import numpy as np
import pandas as pd
from sklearn.model_selection import train_test_split
from sklearn.preprocessing import LabelEncoder
from sklearn.metrics import classification_report
from sklearn.feature_extraction.text import TfidfVectorizer
from tensorflow.keras import models
from tensorflow.keras import layers
from keras.callbacks import ModelCheckpoint
import pickle
import re
import nltk
nltk.download('stopwords')
nltk.download('wordnet')

# leitura dos dados
df = pd.read_table('./data/train.txt', header=None, names=['text', 'target'])

# aplicação da remoção das stopwords, stemming
def preprocess_text(text):
    stopwords = nltk.corpus.stopwords.words("english")
    ## clean (convert to lowercase and remove punctuations and characters and then strip)
    text = re.sub(r'[^\w\s]', '', str(text).lower().strip())
            
    ## Tokenize (convert from string to list)
    text_ = text.split()
    ## remove Stopwords
    text_ = [word for word in text_ if word not in stopwords]
                
    ## Stemming (remove -ing, -ly, ...)
    ps = nltk.stem.porter.PorterStemmer()
    text_ = [ps.stem(word) for word in text_]
                         
    ## back to string from list
    text = " ".join(text_)
    return text

df['text_clean'] = df['text'].apply(lambda x: preprocess_text(x))

print(df['text_clean'].head(2))


# codificação das labels
le = LabelEncoder()
df['target']  = le.fit_transform(df['target'])

# Divisão em variável target e variável preditora
X = df['text_clean'].values
y = df['target'].values

# Vetorização
tfidf = TfidfVectorizer()
X_tfidf = tfidf.fit_transform(X)

# salvando o objeto tfidf
pickle.dump(tfidf, open("./models/tfidf.pickle", "wb"))

# divisão dos dados
X_train, X_valid, y_train, y_valid = train_test_split(X_tfidf, 
                                                        y, 
                                                        test_size=0.2, 
                                                        random_state=42, 
                                                        stratify=y)

X_train_v2, X_test, y_train_v2, y_test = train_test_split(X_train, 
                                                            y_train, 
                                                            test_size=0.25, 
                                                            random_state=42, 
                                                            stratify=y_train)

path_best_model = './models/model.weights.best.hdf5'
checkpointer = ModelCheckpoint(filepath=path_best_model, 
                               verbose=1,
                              save_best_only=True)

# criação do modelo
model = models.Sequential()
model.add(layers.Dense(64, activation='relu', input_shape=(20352, )))
model.add(layers.Dense(64, activation='relu'))
model.add(layers.Dense(3, activation='sigmoid'))

model.compile(optimizer='rmsprop', 
              loss='sparse_categorical_crossentropy', 
              metrics=['accuracy'])

# treino
history = model.fit(X_train_v2.toarray(), 
                    y_train_v2, 
                    epochs=50, 
                    batch_size=16, 
                    callbacks=[checkpointer],
                    validation_data=(X_valid.toarray(), y_valid))

# carregando o melhor modelo
model.load_weights(path_best_model)

#  avaliação
loss, accuracy = model.evaluate(X_test.toarray(), y_test)

y_pred = model.predict(X_test.toarray())
y_pred = [ np.argmax(y) for y in y_pred]

print(classification_report(y_test, y_pred, target_names=le.classes_))

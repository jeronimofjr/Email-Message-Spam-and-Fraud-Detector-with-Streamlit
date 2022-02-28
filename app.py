import pandas as pd
import numpy as np
import streamlit as st
from keras.models import load_model

tfidf = pd.read_pickle('./models/tfidf.pickle')
model = load_model('./models/model.weights.best.hdf5')

def prediction(text):
    pred = model.predict(text)
    return pred

def pre_process(text):
    return tfidf.transform([text]).toarray()
   
def get_class(value):
    if value == 0:
        return 'FRAUD'
    elif value == 1:
        return 'NORMAL'
    else:
        return 'SPAM'

def main():
    st.title("Spam or Fraud Message Prediction")
    st.write("This app is created to predict if a email message is Spam, Fraud or Normal")
    text_input = st.text_area('Enter some text')
    result = None
    value = None

    vec = pre_process(text_input)
    if st.button("Predict"):
        value = np.argmax(prediction(vec))
        result = get_class(value) 
        st.subheader('Prediction')
        st.markdown(f'The predicted message is: **{result}**' )
   

if __name__=='__main__':
    main()
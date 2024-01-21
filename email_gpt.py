from openai import OpenAI

client = OpenAI()

import requests
import json
import os
import pinecone
from O365 import Account
from dotenv import load_dotenv
import speech_recognition as sr
import wave
import chromadb
from bs4 import BeautifulSoup
import streamlit as st



load_dotenv()

client_id = os.getenv('CLIENT_ID')
client_secret = os.getenv('CLIENT_SECRET')
pinecone_api = os.getenv('PINECONE_API')



#*******AUDIO FUNCTIONS*********
def record_audio():
    # Load the speech recognizer and set the initial energy threshold and pause threshold
    r = sr.Recognizer()
    r.energy_threshold = 300
    r.pause_threshold = 0.8
    r.dynamic_energy_threshold = False

    with sr.Microphone(sample_rate=16000) as source:
        print("Say something!")
        # Get and save audio to wav file
        audio = r.listen(source)
        
        # Define the file name
        filename = "recording.wav"
        
        # Save the audio data to a file
        with wave.open(filename, 'wb') as wf:
            wf.setnchannels(1)
            wf.setsampwidth(audio.sample_width)
            wf.setframerate(16000)
            wf.writeframes(audio.get_wav_data())

            return filename
        

def transcribe_forever(audio_file_path):
    
    # Start transcription
    with open(audio_file_path, "rb") as audio_file:
        result = client.audio.transcriptions.create(model = "whisper-1", file =  audio_file)
    predicted_text = result.text
    return predicted_text


#**************PINECONE*****************
# Initialize the Pinecone client
pinecone.init(api_key=pinecone_api, environment='us-central1-gcp')

# List the indexes
list_indexes = pinecone.list_indexes()
print("List of indexes:", list_indexes)

# Setup the Pinecone vector index
index_name = "email-testing"
index = pinecone.Index(index_name)
print("Index name:", index_name)

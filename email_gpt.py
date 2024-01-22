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


#*****CHROMA DB*************#

'''
client_chromaDB = chromadb.PersistentClient(path="")

def createChatHistoryEmbeddings(chat_history):
    chat_text = '\n'.join(chat_history)
    embeddings = client.embeddings.create(
        input=chat_text,
        model="text-embedding-ada-002"
    )
    # Extract the embedding
    embedding_chatHistory = embeddings.data[0].embedding
    return embedding_chatHistory


def store_embeddings_in_chroma_db(embedding_chatHistory, client_chromaDB, collection_name):
    collection = client_chromaDB.get_or_create_collection(collection_name)
    collection.add(
        embeddings=embedding_chatHistory,
        #metadatas=metadata,
        ids=str(your_unique_identifier_here)  # Unique identifier for the chat entry
    )
'''


#*************INBOX*****************
#Outlook Graph API credentials
credentials = (client_id, client_secret)
account = Account(credentials)

if account.authenticate(scopes=['basic', 'message_all']):
    print('Authenticated!')


def retrieveMessages():
    # Access the mailbox
    mailbox = account.mailbox()
    inbox = mailbox.inbox_folder()
    messages_ = []  # Initialize an empty list to store messages

# Get unread messages
    unread_messages = inbox.get_messages(limit=25, query="isRead eq false")



    for message in unread_messages:
         # Parse HTML and extract text and/or tables
        #soup = BeautifulSoup(message.body, 'html.parser')
        #text_and_tables = ''.join([str(tag) for tag in soup.find_all(['p', 'table'])])
    
        attachments_info = []
        if message.has_attachments:
            # Download attachments into memory (not saved to disk)
            message.attachments.download_attachments()

            for attachment in message.attachments:
                
                attachments_info.append(attachment.name)

        # Message metadata
        metadata = {
            "ID": message.object_id,
            "Conversation ID": message.conversation_id,
            "Subject": message.subject,
            "Received": message.received,
            #"Body preview": message.body_preview,
            "Attachments": attachments_info if attachments_info else None
        }
        
        # Append the tuple (message body, metadata, id) to the list
        messages_.append((message.body, metadata, message.conversation_id))

    return messages_

# Assuming 'account' is already authenticated
messageData = retrieveMessages()

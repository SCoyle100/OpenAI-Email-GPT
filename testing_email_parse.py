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
import threading



load_dotenv()

client_id = os.getenv('CLIENT_ID')
client_secret = os.getenv('CLIENT_SECRET')
pinecone_api = os.getenv('PINECONE_API')






class AudioRecorder:
    def __init__(self):
        self.recognizer = sr.Recognizer()
        self.microphone = sr.Microphone(sample_rate=16000)
        self.recording = False
        self.thread = None
        self.frames = []

    def start_recording(self):
        self.recording = True
        self.frames = []
        self.thread = threading.Thread(target=self.record)
        self.thread.start()

    def stop_recording(self):
        self.recording = False
        self.thread.join()  # Wait for the recording thread to finish
        filename = "recording.wav"
        with wave.open(filename, 'wb') as wf:
            wf.setnchannels(1)
            wf.setsampwidth(self.recognizer.recognize().sample_width)
            wf.setframerate(16000)
            wf.writeframes(b''.join(self.frames))
        return filename

    def record(self):
        with self.microphone as source:
            self.recognizer.adjust_for_ambient_noise(source)  # Adjust for ambient noise once at the beginning
            while self.recording:
                audio = self.recognizer.listen(source, phrase_time_limit=5)  # Listen for 5 seconds
                self.frames.append(audio.get_wav_data())


#*******AUDIO FUNCTIONS*********
                
'''                
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
'''        

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

def parse_html_content(html_content):
    soup = BeautifulSoup(html_content, 'html.parser')

    # Remove script and style elements
    for script_or_style in soup(["script", "style"]):
        script_or_style.decompose()

    # Handle links
    for link in soup.find_all('a'):
        href = link.get('href', '')
        link_text = link.get_text(strip=True)
        link.replace_with(link_text)

    # Handle images
    for image in soup.find_all('img'):
        alt_text = image.get('alt', '')
        image.replace_with(f'[Image: {alt_text}]' if alt_text else '')

    # Extract text
    text = soup.get_text(separator=' ', strip=True)

    # Replace or remove specific unicode characters
    replacements = {
    '\u200c': '',  # ZERO WIDTH NON-JOINER
    '\xa0': ' ',   # NO-BREAK SPACE
    '͏': '',       # Control character
    '\u200b': '',  # ZERO WIDTH SPACE
    '\u200e': '',  '\u200f': '',  # Directionality marks
    '\u00ad': '',  # SOFT HYPHEN
    '\u200a': ' ', '\u2009': ' ', '\u2002': ' ', '\u2003': ' ',  # Various spaces
    '\u200d': '',  # ZERO WIDTH JOINER
    '\ufffc': '',  # OBJECT REPLACEMENT CHARACTER
    '\u2028': '\n', '\u2029': '\n\n',  # Line and paragraph separators
    '\u2060': '',  # WORD JOINER
    '\u2011': '-'  # NON-BREAKING HYPHEN
}
    
    for search, replace in replacements.items():
        text = text.replace(search, replace)

    # Clean up whitespace further
    lines = (line.strip() for line in text.splitlines())
    chunks = (phrase.strip() for line in lines for phrase in line.split("  "))
    cleaned_text = '\n'.join(chunk for chunk in chunks if chunk)

    return cleaned_text


def retrieveMessages():
    # Access the mailbox
    mailbox = account.mailbox()
    inbox = mailbox.inbox_folder()
    messages_ = []  # Initialize an empty list to store messages

# Get unread messages
    unread_messages = inbox.get_messages(limit=3, query="isRead eq true")



    for message in unread_messages:
         # Parse HTML and extract text and/or tables
        cleaned_email = parse_html_content(message.body)
    
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
        messages_.append((cleaned_email, metadata, message.conversation_id))

    return messages_


        

# Assuming 'account' is already authenticated
messageData = retrieveMessages()

print(messageData)
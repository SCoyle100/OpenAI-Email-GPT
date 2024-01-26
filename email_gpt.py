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


#**********FUNCTIONS FOR VECTOR SEARCH AND QUERY***************
def processMessageData(messageData):
    for message in messageData:
        # Extract the email body, metadata, and ID from each message tuple
        email_body, metadata, message_id = message

        #print("Type:", type(messageData[0]))
        #print("Content:", messageData[0])


        # Create embeddings for the email body
        embedding = createEmbeddings(email_body, metadata)

        # Store the embedding and metadata in the Pinecone vector database
        storeResponse = storeVector(embedding, message_id, metadata)

        # Optionally, you can print the response or handle it as needed
        print(f"Stored message ID {message_id} with response: {storeResponse}")


'''
def createEmbeddings(email_body):
    embeddings = client.embeddings.create(
        input=email_body,
        model="text-embedding-ada-002"
    )
    # Extract the embedding
    embedding = embeddings.data[0].embedding
    return embedding
'''

def createEmbeddings(email_body, metadata):
    # Convert metadata to a text format
    metadata_text = []
    for key, value in metadata.items():
        value_str = str(value) if value is not None else "none"
        metadata_text.append(f"{key}: {value_str}")

    # Combine the metadata and email body
    combined_text = "\n".join(metadata_text) + "\n\n" + email_body

    # Generate embedding for the combined text
    embeddings = client.embeddings.create(
        input=combined_text,
        model="text-embedding-ada-002"
    )

    # Extract the embedding
    embedding = embeddings.data[0].embedding
    return embedding
  


def storeVector(embedding, id, metadata):
    # Convert None values to a string representation
    for key in metadata:
        if metadata[key] is None:
            metadata[key] = "none"


    upsertRequest = {
        'vectors': [
            {
                'id': id,
                'values': embedding,
                'metadata': metadata,
            }
        ]
    }

    upsertResponse = index.upsert(vectors=upsertRequest['vectors'])

    return upsertResponse







def get_index_stats(index):
    index_stats = index.describe_index_stats({})
    print(index_stats)
    return index_stats






def createQueryVector(userQuery):
    queryEmbeddings = client.embeddings.create (
        input=userQuery,
        model="text-embedding-ada-002",
    )

    queryEmbedding = queryEmbeddings.data[0].embedding
    #print(queryEmbedding)
    return queryEmbedding


def contextSearch(queryEmbedding):
    searchResponse = index.query(
        vector=queryEmbedding,
        top_k=3,
        includeMetadata=True
    )

    matches = searchResponse.get('matches', [])
    contexts = [
        f"Subject: {result['metadata'].get('subject')}\nDate: {result['metadata'].get('date')}\nBody: {result['metadata'].get('body')}\nID: {result['metadata'].get('ID')}\nAttachments: {result['metadata'].get('Attachments')}" 
        for result in matches 
        if 'body' in result.get('metadata', {})
    ]
    return contexts



def create_prompt(userQuery, queryEmbedding):
    # Get contexts using the contextSearch function
    contexts = contextSearch(queryEmbedding)

    prompt_start = (
        "Answer the question based on the context below. \n\n"
        "Context: \n"
    )

    # Join contexts with line breaks
    joined_contexts = "\n\n".join(contexts)

    prompt_end = f"\nQuestion: {userQuery} \n"

    # Create the final prompt
    query_with_contexts = prompt_start + joined_contexts + prompt_end

    #print(query_with_contexts)

    return query_with_contexts


def sendPrompt(prompt):
    completion = client.chat.completions.create(
        model="gpt-4-1106-preview",
        messages=[{"role": "system", "content": "You are a helpful assistant that answers questions about emails and more. Don't make assumptions about what values to plug into functions. Ask for clarification if a user request is ambiguous."}, 
                  {"role": "user", "content": prompt}]
        )
    
    return completion.choices[0].message.content


messageData = retrieveMessages()

processMessageData(messageData)


# Static variables
main_folder_path = r"C:\Users\seanc\OneDrive\Desktop\Function Calling Main Folder"



#********CALLABLE FUNCTIONS************
def create_folder(folder_name):
    subfolder_path = os.path.join(main_folder_path, folder_name)

    # Check if the folder already exists
    if os.path.exists(subfolder_path):
        return f"Folder already exists: {subfolder_path}"

    try:
        os.makedirs(subfolder_path, exist_ok=True)
        return f"Folder created: {subfolder_path}"
    except OSError as error:
        return f"Error creating folder: {error}"
    

def downloading_attachments(folder_name, message_id):
    folder_path = os.path.join(main_folder_path, folder_name)

    # Ensure the target folder exists or create it if it doesn't
    if not os.path.exists(folder_path):
        os.makedirs(folder_path, exist_ok=True)


        message_id.attachments.download_attachments()

        for attachment in message_id.attachments:
            if not attachment.is_inline:
                # Specify the directory where you want to save the attachments
                #save_directory = r'C:\Users\seanc\OneDrive\Desktop\Programming\Email_AI_Assistant'
                #file_path = os.path.join(save_directory, attachment.name)

                # Save the attachment to the specified directory
                attachment.save(folder_path)

                #attachment.save(file_path)
                print(f"Attachment {attachment.name} saved to {folder_path}")

    return f"Attachments downloaded to {folder_path}"
    

# Function to send an email (mock implementation)
def send_email(person_name):
    print(f"Send email response: {person_name}")


def email_standard_search(query):
    """
    Searches the Outlook inbox for messages matching the given query.

    :param account: Authenticated O365 Account object.
    :param query: The query to search for in the inbox.
    """

    # Ensure the account is authenticated
    if not account.is_authenticated:
        raise ValueError("Account is not authenticated.")

    # Access the mailbox
    mailbox = account.mailbox()
    inbox = mailbox.inbox_folder()

    query = f"(contains(subject, {query}) or contains(body, {query}))"


    # Search the inbox
    messages = inbox.get_messages(limit=25, query=query)

    # Process the messages
    for message in messages:
        print(f"Subject: {message.subject}, Sender: {message.sender}, Received: {message.received}")

# Example usage
# Assuming 'account' is an authenticated O365 Account object
# search_outlook_inbox(account, 'subject:"Important"')



def email_vector_search(userQuery):
    # Create the query embedding using the same encode model as in `createEmbeddings`
    queryEmbedding = createQueryVector(userQuery)

    # Retrieve relevant contexts from the vector database
    contexts = contextSearch(queryEmbedding)

    # Generate a prompt with the additional contexts to send to the completion model
    prompt = create_prompt(userQuery, queryEmbedding)

    # Generate a response using the modified prompt
    answer = sendPrompt(prompt)
    
    print(answer)

    chat_history.append(f"Agent: {answer}")

    return answer
        







available_functions = {
    "create_folder": create_folder,
    "send_email": send_email, 
    "download_attachments": downloading_attachments, #download email attachments
    "email_vector_search": email_vector_search, #query email inbox
    "email_standard_search": email_standard_search
}
from openai import OpenAI

client = OpenAI()


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
    '\u200c': ' ',
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
    unread_messages = inbox.get_messages(limit=4, query="isRead eq false")



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
            "Body preview": message.body_preview,
            "Body": cleaned_email,
            "Attachments": attachments_info if attachments_info else None
        }
        
        # Append the tuple (message body, metadata, id) to the list
        messages_.append((cleaned_email, metadata, message.conversation_id))

    return messages_


        

# Assuming 'account' is already authenticated
messageData = retrieveMessages()

print(messageData)




#**********FUNCTIONS FOR VECTOR SEARCH AND QUERY***************
def processMessageData(messageData):                  #try replacing cleaned_email with email_body again?
    for message in messageData:
        # Extract the email body, metadata, and ID from each message tuple
        cleaned_email, metadata, message_id = message

        #print("Type:", type(messageData[0]))
        #print("Content:", messageData[0])


        # Create embeddings for the email body
        embedding = createEmbeddings(cleaned_email, metadata)

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



def createEmbeddings(cleaned_email, metadata):
    # Convert metadata to a text format
    metadata_text = []
    for key, value in metadata.items():
        value_str = str(value) if value is not None else "none"
        metadata_text.append(f"{key}: {value_str}")

    # Combine the metadata and email body
    combined_text = "\n".join(metadata_text) + "\n\n" + cleaned_email

    # Generate embedding for the combined text
    embeddings = client.embeddings.create(
        input=combined_text,
        model="text-embedding-3-small"
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
        model="text-embedding-3-small",
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
    if not matches:
        return ["No matches found."]

    contexts = []
    for result in matches:
        # Initialize an empty list to hold the result strings
        result_strings = []

        # Match the metadata fields and their naming convention exactly as stored in Pinecone
        metadata_fields = {
            'Subject': 'Subject',
            'Received': 'Date',
            'Body preview': 'Body',
            'Body': 'Body',
            'ID': 'ID',
            'Attachments': 'Attachments'
        }

        for field, display_name in metadata_fields.items():
            value = result['metadata'].get(field, 'N/A')  # Use 'N/A' if the field is not present
            result_strings.append(f"{display_name}: {value}")

        # Join all the result strings into a single string for this result
        context = "\n".join(result_strings)
        contexts.append(context)

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
        model="gpt-4-0125-preview",
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


# Define functions for the assistant
functions = [
    {
        "name": "create_folder",
        "description": "Creates a folder with a specified name.",
        "parameters": {
            "type": "object",
            "properties": {
                "folder_name": {
                    "type": "string",
                    "description": "The name of the folder to be created."
                }
            },
            "required": ["folder_name"]
        }
    },
    {
        "name": "send_email",
        "description": "Sends an email to a specified person.",
        "parameters": {
            "type": "object",
            "properties": {
                "person_name": {
                    "type": "string",
                    "description": "The name of the person to send the email to."
                }
            },
            "required": ["person_name"]
        }
    },
    {
        "name": "downloading_attachments",
  "description": "Downloads email attachments to a specified folder.",
  "parameters": {
    "type": "object",
    "properties": {
      "folder_name": {
        "type": "string",
        "description": "The name of the folder where attachments will be downloaded."
      },
      "files": {
        "type": "array",
        "items": {
          "type": "object",
          "properties": {
           
            "file_name": {
              "type": "string"
            },
            
            
          },
          "required": ["file_name"] 
        },
        "description": "The list of attachment objects to download."
      }
    },
    "required": ["folder_name", "files"]
        }
    },
    {
        "name": "email_vector_search",
        "description": '''Processes a user query to search the pinecone vector database that includes the email inbox by
        retrieving relevant contexts from the pinecone vector database.''',
        "parameters": {
            "type": "object",
            "properties": {
                "userQuery": {
                    "type": "string",
                    "description": "The user query to the email inbox to be processed. this is in the form of a question"
                }
            },
            "required": ["userQuery"]
        }
    },
    {
    "name": "email_standard_search",
    "description": "Searches the Outlook inbox for messages matching a given query. The function accesses the authenticated user's Outlook inbox and retrieves a list of messages that match the search criteria.",
    "parameters": {
        "type": "object",
        "properties": {
            
            "query": {
                "type": "string",
                "description": "The search query to be used for filtering inbox messages."
            }
        },
        "required": ["query"]
    }
}

]




def get_gpt_response(user_input):
    messages = [{"role": "system", "content": "You are a helpful assistant. Don't make assumptions about what values to plug into functions. Ask for clarification if a user request is ambiguous, especially for email search. you may need to confirm if the user wants vector or standard email search."}, {"role": "user", "content": user_input}]
    return client.chat.completions.create(
        model="gpt-4-0125-preview",
        messages=messages,
        tools=[{"type": "function", "function": func} for func in functions],
        tool_choice="auto",
    )

def execute_function_call(function_name, arguments):
    function = available_functions.get(function_name, None)
    if function:
        results = function(**arguments)
    else:
        results = f"Error: function {function_name} does not exist"
    return results

chat_history = []
predefined_prompt = "Here is the chat history for context: "





audio_recorder = AudioRecorder()

while True:
    # Offer the user a choice between typing or speaking
    user_choice = input("Type your message, or press ENTER to start talking and ENTER again when you're done: ")

    if user_choice == "":
        print("Recording... Press ENTER to stop.")
        audio_recorder.start_recording()
        input()  # Wait for the user to press Enter to stop
        audio_file_path = audio_recorder.stop_recording()
        print("Transcribing...")
        user_input = transcribe_forever(audio_file_path)
    else:
        user_input = user_choice  # Use the typed input directly

    # This check allows the user to quit after their last message is transcribed or typed
    if user_input.lower().strip() == "quit":
        print("Quitting...")
        break

    print(f"You said: {user_input}")  # Print user input to the terminal
    chat_history.append(f"User: {user_input}")  # Add user input to chat history

    # Incorporate the chat history into the GPT response if applicable
    if len(chat_history) > 1:
        user_input = predefined_prompt + '\n'.join(chat_history) + "\n" + user_input

    gpt_response = get_gpt_response(user_input)
    gpt_text_response = gpt_response.choices[0].message.content
    print(f"GPT Response: {gpt_text_response}")  # Print GPT response text output to the terminal
    chat_history.append(f"Agent: {gpt_text_response}")  # Add GPT text response to chat history


    # Your existing logic to handle tool calls and errors follows


    # Check for tool_calls in the GPT response
    if hasattr(gpt_response.choices[0].message, 'tool_calls') and gpt_response.choices[0].message.tool_calls:
        tool_call = gpt_response.choices[0].message.tool_calls[0]
        function_name = tool_call.function.name
        arguments = json.loads(tool_call.function.arguments)  # Parse JSON arguments

        try:
            function_response = execute_function_call(function_name, arguments)
            chat_history.append(f"GPT: {function_response}")  # Add function response to chat history
        except Exception as e:
            error_message = f"An error occurred: {str(e)}"
            print(error_message)  # Print the error message to the terminal
            chat_history.append(f"GPT: {error_message}")  # Add error message to chat history
    else:
        function_response = "No function call in response."
        #chat_history.append(f"GPT: {function_response}") 
        #
 
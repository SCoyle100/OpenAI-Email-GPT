from shared_resources import chat_history
import os
import shared_resources




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

    email_processor = shared_resources.email_processor

    if email_processor is None:
        raise ValueError("Error: email_processor is not initialized.")
    """
    Searches the Outlook inbox for messages matching the given query.

    :param account: Authenticated O365 Account object.
    :param query: The query to search for in the inbox.
    """

    # Ensure the account is authenticated
    if not email_processor.account.is_authenticated:
        raise ValueError("Account is not authenticated.")

    # Access the mailbox
    mailbox = email_processor.account.mailbox()
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

    email_processor = shared_resources.email_processor

    if email_processor is None:
        raise ValueError("Error: email_processor is not initialized.")
    
    
    if email_processor is None:
        return "Error: EmailProcessor is not initialized."
    # Create the query embedding using the same encode model as in `createEmbeddings`
    queryEmbedding = email_processor.create_query_vector(userQuery)

    # Retrieve relevant contexts from the vector database
    contexts = email_processor.context_search(queryEmbedding)

    # Generate a prompt with the additional contexts to send to the completion model
    prompt = email_processor.create_prompt(userQuery, queryEmbedding)

    # Generate a response using the modified prompt
    answer = email_processor.send_prompt(prompt)
    
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

from shared_resources import chat_history
import os
import shared_resources


email_processor = shared_resources.email_processor


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
    # Create the query embedding using the same encode model as in `createEmbeddings`
    queryEmbedding = email_processor.createQueryVector(userQuery)

    # Retrieve relevant contexts from the vector database
    contexts = email_processor.contextSearch(queryEmbedding)

    # Generate a prompt with the additional contexts to send to the completion model
    prompt = email_processor.create_prompt(userQuery, queryEmbedding)

    # Generate a response using the modified prompt
    answer = email_processor.sendPrompt(prompt)
    
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
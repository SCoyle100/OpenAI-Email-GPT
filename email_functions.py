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

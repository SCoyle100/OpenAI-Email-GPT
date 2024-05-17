from O365 import Account
from bs4 import BeautifulSoup
import pinecone
import shared_resources



class EmailProcessor:
    def __init__(self, client_id=None, client_secret=None, account=None):
        if account is not None:
            self.account = account
        else:
            self.credentials = (client_id, client_secret)
            self.account = Account(self.credentials)
            self.authenticate()

    def authenticate(self):
        if not self.account.is_authenticated:  # Check if not already authenticated
            if self.account.authenticate(scopes=['basic', 'message_all']):
                print('Authenticated!')
            else:
                raise Exception("Authentication Failed")
        else:
            print("Already authenticated.")
        
    
    def parse_html_content(self, html_content):
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
    

    def retrieveMessages(self):
        mailbox = self.account.mailbox()
        inbox = mailbox.inbox_folder()
        messages_ = []

        # Assuming you want unread messages, change query accordingly
        unread_messages = inbox.get_messages(limit=3, query="isRead eq false")

        for message in unread_messages:
            cleaned_email = self.parse_html_content(message.body)
            attachments_info = []
            if message.has_attachments:
                # Assuming the O365 library handles attachment downloads as described
                message.attachments.download_attachments()
                for attachment in message.attachments:
                    attachments_info.append(attachment.name)

            metadata = {
                "ID": message.object_id,
                "Conversation ID": message.conversation_id,
                "Subject": message.subject,
                "Received": message.received,
                "Body preview": message.body_preview,
                "Body": cleaned_email,
                "Attachments": attachments_info if attachments_info else None
        }
            messages_.append((cleaned_email, metadata, message.conversation_id))

        return messages_

    def retrieve_and_print_messages(self):
        messageData = self.retrieveMessages()
        print(messageData)



class EmailVectorSearch:

    email_processor = shared_resources.email_processor


    def __init__(self, pinecone_api_key, email_processor, index_name="email-testing"):
       

        self.pinecone_api_key = pinecone_api_key
        self.index_name = index_name
        self.email_processor = email_processor  # Use the provided, authenticated EmailProcessor instance
        self.client = client  # Assuming OpenAI() is correctly defined elsewhere
        self.initialize_pinecone()
       

    def initialize_pinecone(self):
        pinecone.init(api_key=self.pinecone_api_key, environment="us-central1-gcp")
        list_indexes = pinecone.list_indexes()
        print("List of Indexes:", list_indexes)

        self.index_name = "email-testing"
        self.index = pinecone.Index(self.index_name)  # Assign the Pinecone index to the class instance
        print("Index name:", self.index)

        '''

    def extractSchema(self):
        # Find the schema for the email_vector_search function
        for function in functions:
            if function['name'] == 'email_vector_search':
                return function['parameters']['properties']['userQuery']
        # Return None or raise an error if the schema is not found
        return None
'''
    

    #This method transfers the messageData from the EmailProcessor to the processMessageData method
    def fetchMessages(self):
        messageData = self.email_processor.retrieveMessages()
        self.processMessageData(messageData)

    

    def processMessageData(self, messageData):
        for message in messageData:
        # Extract the email body, metadata, and ID from each message tuple
            email_body, metadata, message_id = message

            #print("Type:", type(messageData[0]))
            #print("Content:", messageData[0])


        # Create embeddings for the email body
        embedding = self.createEmbeddings(email_body, metadata)

        # Store the embedding and metadata in the Pinecone vector database
        storeResponse = self.storeVector(embedding, message_id, metadata)

        # Optionally, you can print the response or handle it as needed
        print(f"Stored message ID {message_id} with response: {storeResponse}")    

      

    def createEmbeddings(self, cleaned_email, metadata):
        metadata_text = [f"{key}: {str(value) if value is not None else 'none'}" for key, value in metadata.items()]
        combined_text = "\n".join(metadata_text) + "\n\n" + cleaned_email
        embeddings = self.client.embeddings.create(input=combined_text, model="text-embedding-3-small")
        return embeddings.data[0].embedding

    def storeVector(self, embedding, id, metadata):
        for key in metadata:
            if metadata[key] is None:
                metadata[key] = "none"
        upsert_request = {
            'vectors': [
                {
                    'id': id, 
                    'values': embedding, 
                    'metadata': metadata
                }
            ]
        }

        upsert_response = self.index.upsert(vectors=upsert_request['vectors'])

        return upsert_response

    def createQueryVector(self, userQuery):
        queryEmbeddings = self.client.embeddings.create(
            input=userQuery, 
            model="text-embedding-3-small",
            )
        
        queryEmbedding = queryEmbeddings.data[0].embedding
        
        return queryEmbedding
    

    def contextSearch(self, queryEmbedding):
        search_response = self.index.query(
            vector=queryEmbedding, 
            top_k=3, 
            include_metadata=True
            )
        
        matches = search_response.get('matches', [])
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


    def create_prompt(self, userQuery, queryEmbedding):
    # Get contexts using the contextSearch function
        contexts = self.contextSearch(queryEmbedding)

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

    def sendPrompt(self, prompt):
        completion = self.client.chat.completions.create(model="gpt-4-0125-preview", messages=[{"role": "system", "content": "You are a helpful assistant that answers questions about emails and more. Don't make assumptions about what values to plug into functions. Ask for clarification if a user request is ambiguous."}, {"role": "user", "content": prompt}])
        return completion.choices[0].message.content        





functions = [

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
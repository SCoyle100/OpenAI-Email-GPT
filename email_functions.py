
from O365 import Account
from bs4 import BeautifulSoup
import pinecone
import shared_resources
from openai import OpenAI

client = OpenAI()



class EmailProcessor:
    def __init__(self, o365_client_id=None, o365_client_secret=None, o365_account=None, pinecone_api_key=None):
        self.o365_client_id = o365_client_id
        self.o365_client_secret = o365_client_secret
        self.pinecone_api_key = pinecone_api_key
        
        
        if o365_account is not None:
            self.account = o365_account
        else:
            self.credentials = (self.o365_client_id, self.o365_client_secret)
            self.account = Account(self.credentials)
            self.authenticate_o365()
        
        if self.pinecone_api_key:
            self.initialize_pinecone()

    def authenticate_o365(self):
        if not self.account.is_authenticated:
            if self.account.authenticate(scopes=['basic', 'message_all']):
                print('Authenticated with O365!')
            else:
                raise Exception("O365 Authentication Failed")
        else:
            print("Already authenticated with O365.")
    
    def initialize_pinecone(self):
        pinecone.init(api_key=self.pinecone_api_key, environment="us-central1-gcp")
        list_indexes = pinecone.list_indexes()
        print("List of Indexes:", list_indexes)
        
        self.index_name = "email-testing"
        self.index = pinecone.Index(self.index_name)
        print("Index name:", self.index)

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

    def retrieve_messages(self):
        mailbox = self.account.mailbox()
        inbox = mailbox.inbox_folder()
        messages_ = []

        # Assuming you want unread messages, change query accordingly
        unread_messages = inbox.get_messages(limit=3, query="isRead eq false")

        for message in unread_messages:
            cleaned_email = self.parse_html_content(message.body)
            attachments_info = []
            if message.has_attachments:
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
        message_data = self.retrieve_messages()
        print(message_data)

    def fetch_messages(self):
        message_data = self.retrieve_messages()
        self.process_message_data(message_data)

    def process_message_data(self, message_data):
        for message in message_data:
            email_body, metadata, message_id = message
            embedding = self.create_embeddings(email_body, metadata)
            store_response = self.store_vector(embedding, message_id, metadata)
            print(f"Stored message ID {message_id} with response: {store_response}")

    def create_embeddings(self, cleaned_email, metadata):
        metadata_text = [f"{key}: {str(value) if value is not None else 'none'}" for key, value in metadata.items()]
        combined_text = "\n".join(metadata_text) + "\n\n" + cleaned_email
        embeddings = client.embeddings.create(input=combined_text, model="text-embedding-3-small")
        return embeddings.data[0].embedding

    def store_vector(self, embedding, id, metadata):
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

    def create_query_vector(self, user_query):
        query_embeddings = client.embeddings.create(input=user_query, model="text-embedding-3-small")
        query_embedding = query_embeddings.data[0].embedding
        return query_embedding

    def context_search(self, query_embedding):
        search_response = self.index.query(vector=query_embedding, top_k=3, include_metadata=True)
        matches = search_response.get('matches', [])
        if not matches:
            return ["No matches found."]

        contexts = []
        for result in matches:
            result_strings = []

            metadata_fields = {
                'Subject': 'Subject',
                'Received': 'Date',
                'Body preview': 'Body',
                'Body': 'Body',
                'ID': 'ID',
                'Attachments': 'Attachments'
            }

            for field, display_name in metadata_fields.items():
                value = result['metadata'].get(field, 'N/A')
                result_strings.append(f"{display_name}: {value}")

            context = "\n".join(result_strings)
            contexts.append(context)

        return contexts

    def create_prompt(self, user_query, query_embedding):
        contexts = self.context_search(query_embedding)
        prompt_start = (
            "Answer the question based on the context below. \n\n"
            "Context: \n"
        )

        joined_contexts = "\n\n".join(contexts)
        prompt_end = f"\nQuestion: {user_query} \n"
        query_with_contexts = prompt_start + joined_contexts + prompt_end

        return query_with_contexts

    def send_prompt(self, prompt):
        completion = client.chat.completions.create(model="gpt-4-0125-preview", messages=[
            {"role": "system", "content": "You are a helpful assistant that answers questions about emails and more. Don't make assumptions about what values to plug into functions. Ask for clarification if a user request is ambiguous."},
            {"role": "user", "content": prompt}
        ])
        return completion.choices[0].message.content





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

    
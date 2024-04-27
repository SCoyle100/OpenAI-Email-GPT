
from O365 import Account
from bs4 import BeautifulSoup

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
        for script_or_style in soup(["script", "style"]):
            script_or_style.decompose()
        for link in soup.find_all('a'):
            href = link.get('href', '')
            link_text = link.get_text(strip=True)
            link.replace_with(link_text)
        for image in soup.find_all('img'):
            alt_text = image.get('alt', '')
            image.replace_with(f'[Image: {alt_text}]' if alt_text else '')
        text = soup.get_text(separator=' ', strip=True)
        replacements = {
            '\u200c': '', '\xa0': ' ', '͏': '', '\u200b': '', '\u200e': '', '\u200f': '',
            '\u00ad': '', '\u200a': ' ', '\u2009': ' ', '\u2002': ' ', '\u2003': ' ',
            '\u200d': '', '\ufffc': '', '\u2028': '\n', '\u2029': '\n\n', '\u2060': '', '\u2011': '-'
        }
        for search, replace in replacements.items():
            text = text.replace(search, replace)
        lines = (line.strip() for line in text.splitlines())
        chunks = (phrase.strip() for line in lines for phrase in line.split("  "))
        cleaned_text = '\n'.join(chunk for chunk in chunks if chunk)
        return cleaned_text

    def retrieveMessages(self):
        mailbox = self.account.mailbox()
        inbox = mailbox.inbox_folder()
        messages_ = []
        unread_messages = inbox.get_messages(limit=3, query="isRead eq false")
        for message in unread_messages:
            cleaned_email = self.parse_html_content(message.body)
            attachments_info = []
            if message.has_attachments:
                message.attach_attachments()
                for attachment in message.attachments:
                    attachments_info.append(attachment.name)
            metadata = {
                "ID": message.object_id, "Conversation ID": message.conversation_id,
                "Subject": message.subject, "Received": message.received,
                "Body preview": message.body_preview, "Body": cleaned_email,
                "Attachments": attachments_info if attachments_info else None
            }
            messages_.append((cleaned_email, metadata, message.conversation_id))
        return messages_

    def retrieve_and_print_messages(self):
        messageData = self.retrieveMessages()
        print(messageData)

    def processMessageData(self, messageData):
        for message in messageData:
            cleaned_email, metadata, message_id = message
            embedding = self.createEmbeddings(cleaned_email, metadata)
            storeResponse = self.storeVector(embedding, message_id, metadata)
            print(f"Stored message ID {message_id} with response: {storeResponse}")

    def createEmbeddings(self, cleaned_email, metadata):
        metadata_text = [f"{key}: {str(value) if value is not None else 'none'}" for key, value in metadata.items()]
        combined_text = "\n".join(metadata_text) + "\n\n" + cleaned_email
        embeddings = client.embeddings.create(input=combined_text, model="text-embedding-3-small")
        return embeddings.data[0].embedding

    def storeVector(self, embedding, id, metadata):
        metadata = {k: v if v is not None else 'none' for k, v in metadata.items()}
        upsertRequest = {
            'vectors': [{'id': id, 'values': embedding, 'metadata': metadata}]
        }
        upsertResponse = index.upsert(vectors=upsertRequest['vectors'])
        return upsertResponse





class SearchUtility:
    def __init__(self, client, index):
        self.client = client
        self.index = index

    def get_index_stats(self):
        index_stats = self.index.describe_index_stats({})
        print(index_stats)
        return index_stats

    def create_query_vector(self, user_query):
        query_embeddings = self.client.embeddings.create(input=user_query, model="text-embedding-3-small")
        return query_embeddings.data[0].embedding

    def context_search(self, query_embedding):
        search_response = self.index.query(vector=query_embedding, top_k=3, include_metadata=True)
        matches = search_response.get('matches', [])
        if not matches:
            return ["No matches found."]
        contexts = []
        metadata_fields = {'Subject': 'Subject', 'Received': 'Date', 'Body preview': 'Body', 'Body': 'Body', 'ID': 'ID', 'Attachments': 'Attachments'}
        for result in matches:
            result_strings = [f"{display_name}: {result['metadata'].get(field, 'N/A')}" for field, display_name in metadata_fields.items()]
            context = "\n".join(result_strings)
            contexts.append(context)
        return contexts

    def create_prompt(self, user_query):
        query_embedding = self.create_query_vector(user_query)
        contexts = self.context_search(query_embedding)
        prompt_start = "Answer the question based on the context below. \n\nContext: \n"
        joined_contexts = "\n\n".join(contexts)
        prompt_end = f"\nQuestion: {user_query} \n"
        return prompt_start + joined_contexts + prompt_end

    def send_prompt(self, prompt):
        completion = self.client.chat.completions.create(model="gpt-4-0125-preview", messages=[{"role": "system", "content": "You are a helpful assistant that answers questions about emails and more."}, {"role": "user", "content": prompt}])
        return completion.choices[0].message.content








    
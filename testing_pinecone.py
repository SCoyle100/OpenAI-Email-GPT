
import pinecone
import os
from openai import OpenAI

client = OpenAI()


pinecone_api = os.getenv('PINECONE_API')




# Initialize the Pinecone client
pinecone.init(api_key=pinecone_api, environment='us-central1-gcp')

# List the indexes
list_indexes = pinecone.list_indexes()
print("List of indexes:", list_indexes)

# Setup the Pinecone vector index
index_name = "email-testing"
index = pinecone.Index(index_name)
print("Index name:", index_name)



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
        top_k=1,
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






# Main interaction loop
if __name__ == "__main__":
    userQuery = input("Enter your text query:\n")
    query_embedding = createQueryVector(userQuery)
    contexts = contextSearch(query_embedding)
    print("Search Results:")
    for context in contexts:
        print(context)
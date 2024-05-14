import os
import json
from dotenv import load_dotenv
from email_functions import EmailProcessor


from openai import OpenAI
import function_calling

from audio_processing import AudioRecorder
from audio_processing import transcribe_forever

import shared_resources

client = OpenAI()






def main():
    # Load environment variables
    load_dotenv()
    

    client_id = os.getenv('CLIENT_ID')
    client_secret = os.getenv('CLIENT_SECRET')
    pinecone_api_key = os.getenv('PINECONE_API')

    # Instantiate the EmailProcessor
    shared_resources.email_processor = EmailProcessor(client_id, client_secret, pinecone_api_key=pinecone_api_key)

    shared_resources.email_processor.authenticate()

    # Now you can use email_processor to call methods defined in the EmailProcessor class
    # For example, to retrieve and print messages:
    shared_resources.email_processor.retrieve_and_print_messages()

    
    messageData = shared_resources.email_processor.retrieveMessages()

    shared_resources.email_processor.initialize_pinecone()

    shared_resources.email_processor.processMessageData(messageData)
    

    #Instantiate the EmailVectorSearch class
    #shared_resources.email_vector = EmailVectorSearch(pinecone_api_key, shared_resources.email_processor)


    #shared_resources.email_vector.initialize_pinecone()

    #shared_resources.email_vector.processMessageData(messageData)



    def get_gpt_response(client, user_input): #adding client as argument due to main function
        messages = [{"role": "system", "content": "You are a helpful assistant. Don't make assumptions about what values to plug into functions. Ask for clarification if a user request is ambiguous, especially for email search. you may need to confirm if the user wants vector or standard email search."}, {"role": "user", "content": user_input}]
        return client.chat.completions.create(
        model="gpt-4-0125-preview",
        messages=messages,
        tools=[{"type": "function", "function": func} for func in function_calling.functions],
        tool_choice="auto",
    )

    def execute_function_call(function_name, arguments):
        function = function_calling.available_functions.get(function_name, None)
        if function:
            results = function(**arguments)
        else:
            results = f"Error: function {function_name} does not exist"
        return results

    shared_resources.chat_history = []
    predefined_prompt = "Here is the chat history for context: "






    audio_recorder = AudioRecorder()

    while True:
    # Offer the user a choice between typing or speaking
        user_choice = input("Type your message, or press ENTER to start talking and ENTER again when you're done: ")
        #if user_choice.lower() == "quit":
        #    break

        if user_choice == "":
            print("Recording... Press ENTER to stop.")
            audio_recorder.start_recording()
            input()  # Wait for the user to press Enter to stop
            audio_file_path = audio_recorder.stop_recording()
            print("Transcribing...")
            user_input = transcribe_forever(audio_file_path)
        else:
            user_input = user_choice  # Use the typed input directly



        print(f"You said: {user_input}")  # Print user input to the terminal
        shared_resources.chat_history.append(f"User: {user_input}")  # Add user input to chat history

    # Incorporate the chat history into the GPT response if applicable
        if len(shared_resources.chat_history) > 1:
            user_input = predefined_prompt + '\n'.join(shared_resources.chat_history) + "\n" + user_input

        gpt_response = get_gpt_response(client, user_input)
        gpt_text_response = gpt_response.choices[0].message.content
        print(f"GPT Response: {gpt_text_response}")  # Print GPT response text output to the terminal
        shared_resources.chat_history.append(f"Agent: {gpt_text_response}")  # Add GPT text response to chat history




    # Check for tool_calls in the GPT response
        if hasattr(gpt_response.choices[0].message, 'tool_calls') and gpt_response.choices[0].message.tool_calls:
            tool_call = gpt_response.choices[0].message.tool_calls[0]
            function_name = tool_call.function.name
            arguments = json.loads(tool_call.function.arguments)  # Parse JSON arguments

            try:
                function_response = execute_function_call(function_name, arguments)
                shared_resources.chat_history.append(f"GPT: {function_response}")  # Add function response to chat history
            except Exception as e:
                error_message = f"An error occurred: {str(e)}"
                print(error_message)  # Print the error message to the terminal
                shared_resources.chat_history.append(f"GPT: {error_message}")  # Add error message to chat history
        else:
            function_response = "No function call in response."
        #chat_history.append(f"GPT: {function_response}") 
        #
 


if __name__ == "__main__":
    main()
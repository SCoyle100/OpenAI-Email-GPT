import speech_recognition as sr
import wave
import threading

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

def transcribe_forever(audio_file_path, client): #adding client as an argument due to refactoring with main function
    
    # Start transcription
    with open(audio_file_path, "rb") as audio_file:
        result = client.audio.transcriptions.create(model = "whisper-1", file =  audio_file)
    predicted_text = result.text
    return predicted_text
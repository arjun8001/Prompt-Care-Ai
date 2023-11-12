import sounddevice as sd
import soundfile as sf
import openai

from colorama import Fore, Style, init
from pydub.playback import play
import win32com.client as wincom
import configparser

config = configparser.ConfigParser()
config.read('settings.ini')
import time
speak = wincom.Dispatch("SAPI.SpVoice")

init()

def open_file(filepath):
    with open(filepath, 'r', encoding='utf-8') as infile:
        return infile.read()

api_key = config.get('API', 'key')

conversation1 = []  
chatbot1 = open_file(r'C:\Users\gupta\Desktop\talk-to-chatgpt-main\chatbot1.txt')

def chatgpt(api_key, conversation, chatbot, user_input, temperature=0.9, frequency_penalty=0.2, presence_penalty=0):
    openai.api_key = api_key
    conversation.append({"role": "user","content": user_input})
    messages_input = conversation.copy()
    prompt = [{"role": "system", "content": chatbot}]
    messages_input.insert(0, prompt[0])
    completion = openai.ChatCompletion.create(
        model="gpt-4-1106-preview",
        temperature=temperature,
        frequency_penalty=frequency_penalty,
        presence_penalty=presence_penalty,
        messages=messages_input)
    chat_response = completion['choices'][0]['message']['content']
    conversation.append({"role": "assistant", "content": chat_response})
    return chat_response


def print_colored(agent, text):
    agent_colors = {
        "": Fore.YELLOW,
    }
    color = agent_colors.get(agent, "")
    print(color + f"{agent}: {text}" + Style.RESET_ALL, end="")


def record_and_transcribe(duration=8, fs=44100):
    print('Recording...')
    myrecording = sd.rec(int(duration * fs), samplerate=fs, channels=2)
    sd.wait()
    print('Recording complete.')
    filename = r'C:\Users\gupta\Desktop\talk-to-chatgpt-main\myrecording.wav'
    sf.write(filename, myrecording, fs)
    with open(filename, "rb") as file:
        openai.api_key = api_key
        result = openai.Audio.transcribe("whisper-1", file)
    transcription = result['text']
    return transcription

while True:
    user_message = record_and_transcribe()
    response = chatgpt(api_key, conversation1, chatbot1, user_message)
    print_colored("Matt:", f"{response}\n\n")
    # Word to remove
    word_to_remove = "Matt"

    # Remove the word from the string
    output_response = response.replace(word_to_remove, "")
    speak.Speak(output_response)


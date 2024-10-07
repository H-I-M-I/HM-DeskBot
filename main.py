import speech_recognition as sr
import os
import webbrowser
import wikipedia
import open_ai
from api_configuration import apikey
import datetime
import random
import numpy as np
import pyttsx3
import win32com.client
import smtplib

engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
# print(voices[1].id)
engine.setProperty('voice', voices[1].id)

speaker = win32com.client.Dispatch("SAPI.SpVoice")

def speak(audio):
    engine.say(audio)
    engine.runAndWait()

def wishMe():
    hour = int(datetime.datetime.now().hour)
    if 0 <= hour < 12:
        speak("Good Morning!")
    elif 12 <= hour < 18:
        speak("Good Afternoon!")
    else:
        speak("Good Evening!")

    speak("Hello, I am BaeMax. Please tell me how may I help you")

chatStr = ""
def chat(query):
    global chatStr
    print(chatStr)
    open_ai.api_key = apikey
    chatStr += f"You: {query}\n Cinderella: "
    response = open_ai.Completion.create(
        model="text-davinci-003",
        prompt=chatStr,
        temperature=0.7,
        max_tokens=256,
        top_p=1,
        frequency_penalty=0,
        presence_penalty=0
    )
    speak(response["choices"][0]["text"])
    chatStr += f"{response['choices'][0]['text']}\n"
    return response["choices"][0]["text"]

def ai(prompt):
    open_ai.api_key = apikey
    text = f"OpenAI response for Prompt: {prompt} \n *************************\n\n"

    response = open_ai.Completion.create(
        model="text-davinci-003",
        prompt=prompt,
        temperature=0.7,
        max_tokens=256,
        top_p=1,
        frequency_penalty=0,
        presence_penalty=0
    )
    
    try:
        text += response["choices"][0]["text"]
        if not os.path.exists("Openai"):
            os.mkdir("Openai")
    except Exception as e:
        return "Some Error Occurred. Sorry from Cinderella"

    # print(response["choices"][0]["text"])
    # with open(f"Openai/prompt- {random.randint(1, 2343434356)}", "w") as f:
    with open(f"Openai/{''.join(prompt.split('intelligence')[1:]).strip()}.txt", "w") as f:
        f.write(text)

    while 1:
        speak(f"open {text}")

'''
def say(text):
    engine.say(text)
    engine.runAndWait()
    #os.system(f"open {text}")
    #os.startfile(text)
    result = subprocess.run(text, capture_output=True, shell=True)
    # Access the output and return code
    #print(result.stdout)
    #print(f"Return code: {result.returncode}")

'''
def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        # r.pause_threshold =  0.6
        audio = r.listen(source)
        try:
            print("Recognizing...")
            query = r.recognize_google(audio, language="en-US")
            print(f"Himi said: {query}")
            return query
        except Exception as e:
            return "Some Error Occurred. Sorry from Cinderella"

def sendEmail(to, content):
    server = smtplib.SMTP('smtp.gmail.com', 535)
    server.ehlo()
    server.starttls()
    server.login('desktopassistant08@gmail.com', '@123456_')
    server.sendmail('desktopassistant08@gmail.com', to, content)
    server.close()

if __name__ == '__main__':
    wishMe()
    print("Welcome to Cinderella's AI")

    while True:
        print("Listening...")
        query = takeCommand().lower()

        sites = [["youtube", "https://www.youtube.com"], ["wikipedia", "https://www.wikipedia.com"],
                 ["google", "https://www.google.com"],  ["khulna university", "https://ku.ac.bd/"]]

        for site in sites:
            if f"Open {site[0]}".lower() in query.lower():
                speak(f"Opening {site[0]} sir...")
                webbrowser.open(site[1])

        # todo: Add a feature to play a specific song
        if "play music" in query:
            musicPath = r"C:\Users\Music\7966_download_s5_crossing_a_river_ringtone.mp3"

            if os.path.exists(musicPath):
                try:
                    os.startfile(musicPath)
                    print(f"Opened file: {musicPath}")
                except Exception as e:
                    print(f"Failed to open file: {musicPath}")
            else:
                print(f"File not found: {musicPath}")

        elif 'wikipedia' in query:
            speak('Searching Wikipedia...')
            query = query.replace("wikipedia", "")
            results = wikipedia.summary(query, sentences=5)
            speak("According to Wikipedia")
            print(results)
            speak(results)

        elif "the time" in query:
            hour = datetime.datetime.now().strftime("%H")
            min = datetime.datetime.now().strftime("%M")
            speak(f"Sir time is {hour} and {min} minutes")
        
        elif "open stackoverflow" in query:
            webbrowser.open("stackoverflow.com")

        elif "MySQL Workbench".lower() in query.lower():
            file_path = r"D:\OneDrive\Desktop\MySQL Workbench 8.0 CE.lnk"

            if os.path.exists(file_path):
                try:
                    os.startfile(file_path)
                    print(f"Opened file: {file_path}")
                except Exception as e:
                    print(f"Failed to open file: {file_path}")
            else:
                print(f"File not found: {file_path}")
        
        elif 'send an email' in query:
            try:
                speak("What should I say?")
                content = takeCommand()
                to = "desktopassistant08@gmail.com"
                sendEmail(to, content)
                speak("Email has been sent!")
            except Exception as e:
                print(e)
                speak("Sorry! I am not able to send this email")

        elif "Using artificial intelligence".lower() in query.lower():
            ai(prompt=query)

        elif "Cinderella Quit".lower() in query.lower():
            exit()

        elif "reset chat".lower() in query.lower():
            chatStr = ""

        else:
            print("Chatting...")
            chat(query)

        # speak(query)

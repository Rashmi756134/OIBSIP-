import speech_recognition as sr
import win32com.client
import webbrowser
import datetime
import pygame
import openai
import os

def say(text):
    speaker = win32com.client.Dispatch("SAPI.SpVoice")
    speaker.Speak(text)

def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        r.pause_threshold = 1
        audio = r.listen(source)
        try:
            print("Recognizing...")
            query = r.recognize_google(audio, language="en-in")
            print(f"User said :{query}")
            return query
        except Exception as e:
            return "Some Error Occurred. Sorry from Jarvis"

if __name__ == '__main__':
    print("Pycharm")
    say("Hello I am Jarvis AI")

    pygame.mixer.init()

    while True:
        print("Listening...")
        query = takeCommand()
        sites = [["youtube", "https://www.youtube.com"], ["wikipedia", "https://www.wikipedia.com"],
                 ["linkedin", "www.linkedin.com/in/rashmi-ranjan-das-b85538262/"], ["Google", "https://www.google.com"]]
        for site in sites:
            if f"Open {site[0]}".lower() in query.lower():
                say(f"Opening {site[0]} sir...")
                webbrowser.open(site[1])
        if "open music" in query:
            musicPath = r"C:\Users\Rashmi Ranjan Das\Downloads\downfall-21371.mp3"
            pygame.mixer.music.load(musicPath)
            pygame.mixer.music.play()

        if "stop music" in query:
            pygame.mixer.music.stop()


        if "the time" in query:
            hour = datetime.datetime.now().strftime("%H")
            min = datetime.datetime.now().strftime("%M")
            say(f"Sir time is {hour} hour and {min} minutes")

        apps = [["Replit", r"C:\Users\Rashmi Ranjan Das\AppData\Local\replit\Replit.exe"],
                ["Firefox", r"C:\Program Files\Mozilla Firefox\firefox.exe"],
                ["Github", r"C:\Users\Rashmi Ranjan Das\AppData\Local\GitHubDesktop\GitHubDesktop.exe"],
                ["Vs code", r"C:\Users\Rashmi Ranjan Das\AppData\Local\Programs\Microsoft VS Code\Code.exe"]]

        for app in apps:
            if f"open {app[0]}".lower() in query.lower():
                os.startfile(app[1])
                say(f"Opening {app[0]} sir...")

        # Check for searching the web
        if "search for" in query.lower():
            search_query = query.lower().replace("search for", "")
            webbrowser.open(f"https://www.google.com/search?q={search_query}")




import subprocess
import wolframalpha
import pyttsx3
import tkinter
import json
import random
import operator
import speech_recognition as sr
import datetime
import wikipedia
import webbrowser
import os
import openai
# import winshell
import pyjokes
# import feedparser
import smtplib
import ctypes
import time
import requests
import shutil
# from twilio.rest import client
# from clint.textui import progress
from ecapture import ecapture as ec
from bs4 import BeautifulSoup
import win32com.client as wincl
from urllib.request import urlopen

engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[1].id)
assname = "Jarvis 1 point o"


def speak(audio):
    engine.say(audio)
    engine.runAndWait()

# text_to_speak = "Hello, this is an example of text-to-speech synthesis."
# speak(text_to_speak)

def wishMe():
    hour = int(datetime.datetime.now().hour)
    if hour>= 0 and hour<12:
        speak("Good Morning Sir !")
  
    elif hour>= 12 and hour<18:
        speak("Good Afternoon Sir !")   
  
    else:
        speak("Good Evening Sir !")  
  
    assname = "Jarvis 1 point o"
    speak("I am your Assistant")
    speak(assname)
 
def username():
    speak("What should i call you sir")
    uname = takeCommand()
    speak("Welcome Mister")
    speak(uname)
    # columns = shutil.get_terminal_size().columns
     
    print("#####################")
    print("Welcome Mr. ", uname)
    print("#####################")
     
    speak("How can i Help you, Sir")


def takeCommand():
     
    r = sr.Recognizer()
     
    with sr.Microphone() as source:
         
        print("Listening...")
        r.pause_threshold = 1
        audio = r.listen(source)
  
    try:
        print("Recognizing...")    
        query = r.recognize_google(audio, language ='en-in')
        print(f"User said: {query}\n")
  
    except Exception as e:
        print(e)    
        print("Unable to Recognize your voice.")  
        return "None"
     
    return query

# print(takeCommand())
# username()

if __name__ == '__main__':
    clear = lambda: os.system('cls')
     
    # This Function will clean any
    # command before execution of this python file
    clear()
    # wishMe()
    # username()
    
    for i in range (1):
         
        speak("How can i help you sir?")
        query = takeCommand().lower()
         
        # All the commands said by user will be 
        # stored here in 'query' and will be
        # converted to lower case for easily 
        # recognition of command
        
        if 'wikipedia' in query:
            speak('Searching Wikipedia...')
            query = query.replace("wikipedia", "")
            results = wikipedia.summary(query, sentences = 3)
            speak("According to Wikipedia")
            print(results)
            speak(results)

        elif "who" in query or "what" in query or "how" in query or "when" in query:
            appid = "HHKQWH-9RQH5X85LX"
            # speak("What should I search for, sir?")
            # query = takeCommand()

            # Encode the query for the URL
            query_encoded = requests.utils.quote(query)

            # Construct the Wolfram Alpha API URL
            url = f"http://api.wolframalpha.com/v1/result?appid={appid}&i={query_encoded}"

            # Send the request and get the response
            response = requests.get(url)

            # Check if the request was successful (HTTP status code 200)
            if response.status_code == 200:
                answer = response.text
                print("Answer:", answer)
                speak(f"The answer is: {answer}")
            else:
                print("Error: Unable to fetch the answer from Wolfram .")
                speak("Sorry, I couldn't find an answer to your query.")

        elif 'search' in query or 'play' in query:
             
            query = query.replace("search", "") 
            query = query.replace("play", "")          
            webbrowser.open(query) 
        
        elif 'open youtube' in query:
            speak("Here you go to Youtube\n")
            webbrowser.open("youtube.com")
 
        elif 'open google' in query:
            speak("Here you go to Google\n")
            webbrowser.open("google.com")

        elif 'open stackoverflow' in query:
            speak("Here you go to Stack Over flow.Happy coding")
            webbrowser.open("stackoverflow.com")

        elif 'music' in query:
            speak("Here you go ")
            webbrowser.open("spotify.com")

        elif 'the time' in query:
            strTime = datetime.datetime.now().strftime("%H:%M:%S")    
            speak(f"Sir, the time is {strTime}")
 
        elif 'open the office' in query:
            codePath = r"E:\Series\The Office\S03\The.Office.US.S03E12.Traveling.Salesmen.720p.WEBRip.2CH.x265.HEVC-PSA.mkv"
            speak("Enjoy your series sir")
            os.startfile(codePath)
 
        elif 'fine' in query or "good" in query:
            speak("It's good to know that your fine")
 
        elif "change my name to" in query:
            query = query.replace("change my name to", "")
            assname = query
 
        elif "change name" in query:
            speak("What would you like to call me, Sir ")
            assname = takeCommand()
            speak("Thanks for naming me")

        elif 'exit' in query:
            speak("Thanks for giving me your time")
            exit()
 
        elif 'joke' in query:
            joke = pyjokes.get_joke()
            print(joke)
            speak(joke)
    
        elif 'change background' in query:
            wallpapers = [
                "pexels-reynaldo-brigworkz-brigantty-771881.jpg",
                "pexels-pixabay-326055.jpg",
                "vincent-guth-62V7ntlKgL8-unsplash.jpg",
                # Add more file paths as needed
            ]
            selected_wallpaper = random.choice(wallpapers)
            ctypes.windll.user32.SystemParametersInfoW(20, 0, r"C:\Users\hp\Downloads\Wallpapers\\" + selected_wallpaper, 0)
            speak("Background changed successfully")

        elif 'news' in query:
             
            try: 
                speak("Which type of news do you want?")
                topic =  takeCommand()
                topic = topic.split()[0]
                url = f"https://newsapi.org/v2/everything?q={topic}&from=2023-09-22&apiKey=cd4d09945be948b592f08f687f2dc895"
                data = requests.get(url).json()
                speak(f'here are some top {topic} news ')
                i = 1
                if(data['status'] == 'ok'):
                    for item in data['articles']:
                        news = str(i) + '. ' + item['title'] + '\n'
                        print(news)
                        speak(news)
                        print(item['description'] + '\n')
                        i += 1
                        if  i == 6:
                            break
                else :
                    speak("Sorry could not fetch news at the moment")
                    

            except Exception as e:
                 
                print(str(e))

        elif 'lock window' in query:
                speak("locking the device")
                ctypes.windll.user32.LockWorkStation()
                exit()
        
        elif "where is" in query:
            query = query.replace("where is", "")
            location = query
            speak("User asked to Locate")
            speak(location)
            webbrowser.open("https://www.google.nl/maps/place/" + location + "")

        elif "write a note" in query:
            try:
                speak("What should I write, sir")
                note = takeCommand()
                file = open('D:\Projects\Voice Assistant/jarvis.txt', 'a')
                strTime = datetime.datetime.now().strftime("%H:%M:%S")
                file.write('\n' + strTime)
                file.write(" :- ")
                file.write(note)
                speak("Note added successfully")
            except Exception as e:
                speak(f"An error occurred: {str(e)}")

        elif "show note" in query:
            speak("Showing Notes")
            file = open("jarvis.txt", "r") 
            print(file.read())
            speak(file.read(6))

        elif "weather" in query:
            api_key = "a0f2b107ccef0893511736884ad593fa"
            base_url = "http://api.openweathermap.org/data/2.5/weather?"
            speak("Which city would you like to know the weather of?")
            print("City name: ")
            city_name = takeCommand()
            complete_url = base_url + "appid=" + api_key + "&q=" + city_name
            response = requests.get(complete_url)
            data = response.json()

            if data["cod"] == 200:
                main_data = data["main"]
                current_temperature = main_data["temp"]
                current_pressure = main_data["pressure"]
                current_humidity = main_data["humidity"]
                weather_data = data["weather"][0]
                weather_description = weather_data["description"]
                forecast = "Temperature  = " + str(current_temperature - 273)  + " degree celsius"+ "\nAtmospheric pressure (in hPa unit) = " + str(current_pressure) + "\nHumidity (in percentage) = " + str(current_humidity) + "\nDescription = " + str(weather_description)
                print(forecast)
                speak(forecast)
            else:
                speak("City Not Found")



        



        


        


 



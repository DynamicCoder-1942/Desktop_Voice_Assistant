import datetime
import time
import speech_recognition as sr
import wikipedia
import webbrowser
import json

import requests

from playsound import playsound


def speak(str):
    from win32com.client import Dispatch

    speaking = Dispatch("SAPI.SpVoice")

    speaking.speak(str)


def wishing_user():
    hour = int(datetime.datetime.now().hour)

    if 0 <= hour < 12:
        speak("Good Morning, Sir. How can I help you")
    elif 12 <= hour <= 18:
        speak("Good Afternoon, Sir.  How can I help you")
    else:
        speak("Good Evening, Sir.  How can I help you")


def user_input():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        r.pause_threshold = 1
        audio = r.listen(source)

    try:
        print("Recognizing...")
        query = r.recognize_google(audio, language='en-in')
        print(f"User said: {query}\n")

    except Exception as e:
        print("Say that again please")

        return "None"
    return query


if __name__ == '__main__':
    playsound('starting up.mp3')
    time.sleep(1)
    speak("powering up")
    time.sleep(2)
    wishing_user()

    while True:
        query = user_input().lower()

        if 'wikipedia' in query:
            speak("Searching Wikipedia ...")
            query = query.replace("wikipedia", "")
            results = wikipedia.summary(query, sentences=3)
            speak("According to the Wikipedia")
            print(results)
            speak(results)

        elif 'open youtube' in query:
            webbrowser.open("youtube.com")

        elif 'open facebook' in query:
            webbrowser.open("facebook.com")

        elif 'open google' in query:
            webbrowser.open("google.com")

        elif 'play music' in query:
            webbrowser.open(
                "https://www.youtube.com/results?search_query=top+music")

        elif 'the time' in query:
            time_now = datetime.datetime.now().strftime("%H:%M:%S")
            print(time_now)
            speak(f"now the time is {time_now}")

        elif 'online compiler' in query:
            webbrowser.open("https://www.onlinegdb.com/")

        elif 'college website' in query:
            webbrowser.open("https://dypiemr.collpoll.com/#/feed")

        elif 'open stack overflow' in query:
            webbrowser.open("stackoverflow.com")

        elif 'news' in query:

            user_input = int(input("Select the category of news u would like to hear from below : \n"
                                   "1. Top Headlines\n"
                                   "2. Entertainment\n"
                                   "3. Health\n"
                                   "4. Science\n"
                                   "5. Sport\n"
                                   "6. Technology\n"
                                   "---->>>> "))

            if user_input == 1:

                news = requests.get(
                    "https://newsapi.org/v2/top-headlines?country=in&category=business&apiKey=e12a2dc1880047ea8262ca31471417ba").text
                news_json = json.loads(news)

                read = news_json['articles']

                for article in read:
                    print(article['title'])
                    speak(article['title'])
                    speak("Moving to the next news")

            elif user_input == 2:
                news = requests.get(
                    "https://newsapi.org/v2/top-headlines?country=in&category=entertainment&apiKey"
                    "=e12a2dc1880047ea8262ca31471417ba").text
                news_json = json.loads(news)

                read = news_json['articles']

                for article in read:
                    print(article['title'])
                    speak(article['title'])
                    speak("Moving to the next news")

            elif user_input == 3:
                news = requests.get(
                    "https://newsapi.org/v2/top-headlines?country=in&category=health&apiKey=e12a2dc1880047ea8262ca31471417ba").text
                news_json = json.loads(news)

                read = news_json['articles']

                for article in read:
                    print(article['title'])
                    speak(article['title'])
                    speak("Moving to the next news")

            elif user_input == 4:
                news = requests.get(
                    "https://newsapi.org/v2/top-headlines?country=in&category=science&apiKey=e12a2dc1880047ea8262ca31471417ba").text
                news_json = json.loads(news)

                read = news_json['articles']

                for article in read:
                    print(article['title'])
                    speak(article['title'])
                    speak("Moving to the next news")

            elif user_input == 5:
                print(article['title'])
                news = requests.get(
                    "https://newsapi.org/v2/top-headlines?country=in&category=sports&apiKey=e12a2dc1880047ea8262ca31471417ba").text
                news_json = json.loads(news)

                read = news_json['articles']

                for article in read:
                    print(article['title'])
                    speak(article['title'])
                    speak("Moving to the next news")

            elif user_input == 6:
                news = requests.get(
                    "https://newsapi.org/v2/top-headlines?country=in&category=technology&apiKey"
                    "=e12a2dc1880047ea8262ca31471417ba").text
                news_json = json.loads(news)

                read = news_json['articles']

                for article in read:
                    print(article['title'])
                    speak(article['title'])
                    speak("Moving to the next news")

        elif 'games' in query:

            print("_____GUESSING THE NUMBER GAME_____")

            import random

            rand1_int = random.randint(0, 100)

            print(":: Rules for playing the Game ::\n"
                  "1. You have to Guess the number between 0 to 100\n"
                  "2. You have only 9 chances to Guess the right number")
            print("\n"
                  "\n"
                  "\n")

            name = str(input("Enter Your Name to begin ----> "))
            namep = name.capitalize()

            print("START TO GUESS YOU NUMBER")

            i = 0
            while (i <= 8):
                user_input = int(input("Start Guessing :: "))

                if rand1_int < user_input:
                    print("The Number you have chosen is greater than the number to be guessed.\n"
                          "TRY AGAIN ")
                # print("Number of GUESSES left :: ", 9-i)

                elif rand1_int > user_input:
                    print("The Number you have chosen is less than the number to be guessed.\n"
                          "TRY AGAIN")
                    # print("Number of GUESSES left :: ", 9-i)

                elif rand1_int == user_input:
                    print("HURRAY....!!!!\n",
                          namep, " YOU HAVE GUESSED THE RIGHT NUMBER")
                    print("Number of Guesses taken ::", i + 1)

                if rand1_int == user_input:
                    print("___WINNER___\n"
                          "You have guessed the number in", i, " chances")
                    break
                print("Number of Guesses left ::", (8 - i))

                i = i + 1

            if i == 9:
                print("___YOU LOST___\n", namep, "Try Again")
                break

        elif 'exit' in query:
            speak(
                "thank you for using this voice assistant.hope to see you seen. good bye sir")
            playsound('exiting music.mp3')
            exit()

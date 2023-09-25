from win32com.client import Dispatch
import requests
import json
import time


def speak(str):
    speaker = Dispatch("SAPI.Spvoice")
    speaker.Speak(str)


if __name__ == "__main__":
    def news_reader():
        print("Hello! Welcome to the world of news. Enter the topic you want to hear about.\n"
              "Ex - Politics, Business, sports, science, technology, health, entertainment")
        speak(
            "  Hello! Welcome to the world of news. Enter the topic you want to hear about.")
        search = input()
        url = f"https://newsapi.org/v2/top-headlines?country=in&category={search}&apiKey=eb532a835aa84885ab7110e7696a594e"
        newwws = requests.get(url).text
        newws = json.loads(newwws)
        news = newws["articles"]
        countt = int(
            input("\nEnter the number of headlines you want to listen to?\n"))
        print()
        count = 1
        i = 0

        time.sleep(0.25)
        while count <= countt:

            if count == 1:
                speak("  The first news for today is  ")

            elif count < countt and count != 1:
                speak(" Moving on to our next news ")

            elif count == countt:
                speak(" So our last news for today is ")

            print(news[i]["title"])
            print(f"For more info visit - {news[i]['url']}\n")
            speak(news[i]["title"])
            time.sleep(0.5)
            count = count + 1
            i = i + 1

        print("Thank you.\n")
        speak("  Thank you.")
        print("Credits - VAIBHAV(me)\nFor any query or complaint contact me.\n")
    news_reader()
    time.sleep(1.5)
    while True:
        print("Do you want to rerun the program?\nType 'yes' or 'no'")
        speak("  Do you want to rerun the program? Type 'yes' or 'no'")
        yn_check = input().capitalize()
        print()
        if yn_check == "Yes":
            news_reader()
            time.sleep(1.5)
            continue
        elif yn_check == "No":
            exit()
        else:
            print("Invalid input. Please type 'yes' or 'no'")
            continue

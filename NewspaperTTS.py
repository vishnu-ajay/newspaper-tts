import requests
import json


def speak(str):
    from win32com.client import Dispatch
    speak = Dispatch("SAPI.Spvoice")
    speak.Speak(str)


def news(str):
    speak("Loading news...")

    url = ('https://newsapi.org/v2/top-headlines?'
           f'country=in&category={str.lower()}&'
           'apiKey=07b7baa694bf4b48b29fb9be3d1ab112')

    respose = requests.get(url)
    text = respose.text
    my_json = json.loads(text)

    speak(f"Top 5 {str} News are")
    for i in range(0, 6):
        speak(my_json['articles'][i]['title'])


if __name__ == '__main__':
    speak("Welcome! Which type of news would you like to hear?")
    speak("Press 1 for Business")
    speak("Press 2 for Entertainment")
    speak("Press 3 for Health")
    speak("Press 4 for Science")
    speak("Press 5 for Sports")
    speak("Press 6 for Technology")

x = int(input("Enter key: "))

if x == 1:
    news('Business')
elif x == 2:
    news('Entertainment')
elif x == 3:
    news('Health')
elif x == 4:
    news('Science')
elif x == 5:
    news('Sports')
elif x == 6:
    news('Technology')
else:
    speak("Wrong input")

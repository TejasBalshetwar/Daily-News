from win32com.client import Dispatch
import requests
import json

def speak(str):
    speak=Dispatch("SAPI.Spvoice")
    speak.speak(str)
if __name__=='__main__':
    k=int(input(" 1 for Business \n 2 for Entertainment \n 3 for Health\n 4 for Science \n 5 for Sports \n 6 for Technology\n"))
    if k==1:
        speak("News for Today....Let's begin")
        url="http://newsapi.org/v2/top-headlines?country=in&category=business&apiKey=055ec88eabc444a98bac843600a4df9b"
        news=requests.get(url).text
        news1=json.loads(news)
        arts=news1['articles']
        for article in arts:
            speak(article['title'])
            speak("Moving on to the next news")
        speak("Thanks for listening")
    elif k==2:
        speak("News for Today....Let's begin")
        url="http://newsapi.org/v2/top-headlines?country=in&category=entertainment&apiKey=055ec88eabc444a98bac843600a4df9b"
        news=requests.get(url).text
        news1=json.loads(news)
        arts=news1['articles']
        for article in arts:
            speak(article['title'])
            speak("Moving on to the next news")
        speak("Thanks for listening")
    elif k==3:
        speak("News for Today....Let's begin")
        url="http://newsapi.org/v2/top-headlines?country=in&category=health&apiKey=055ec88eabc444a98bac843600a4df9b"
        news=requests.get(url).text
        news1=json.loads(news)
        arts=news1['articles']
        for article in arts:
            speak(article['title'])
            speak("Moving on to the next news")
        speak("Thanks for listening")
    elif k==4:
        speak("News for Today....Let's begin")
        url="http://newsapi.org/v2/top-headlines?country=in&category=science&apiKey=055ec88eabc444a98bac843600a4df9b"
        news=requests.get(url).text
        news1=json.loads(news)
        arts=news1['articles']
        for article in arts:
            speak(article['title'])
            speak("Moving on to the next news")
            speak("News for Today....Let's begin")
        speak("Thanks for listening")
        
    elif k==5:
        speak("News for Today....Let's begin")
        url="http://newsapi.org/v2/top-headlines?country=in&category=sports&apiKey=055ec88eabc444a98bac843600a4df9b"
        news=requests.get(url).text
        news1=json.loads(news)
        arts=news1['articles']
        for article in arts:
            speak(article['title'])
            speak("Moving on to the next news")
        speak("Thanks for listening")
    elif k==6:
        speak("News for Today....Let's begin")
        url="http://newsapi.org/v2/top-headlines?country=in&category=technology&apiKey=055ec88eabc444a98bac843600a4df9b"
        news=requests.get(url).text
        news1=json.loads(news)
        arts=news1['articles']
        for article in arts:
            speak(article['title'])
            speak("Moving on to the next news")
        speak("Thanks for listening")
    else:
        print("Enter valid number !!!!!")

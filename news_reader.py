from win32com.client import Dispatch
import requests
import  json
def speak(str):
    speak=Dispatch("SAPI.SpVoice")
    speak.Speak(str)

if __name__ == '__main__':
    speak("Welcome everyone....topmost 10 news of India")
    url="https://newsapi.org/v2/top-headlines?country=in&apiKey=bc6d6a22161448608d247795da4fdfec"
    news=requests.get(url).text
    news_dict=json.loads(news)
    arts=news_dict['articles']
    for article in arts:
        speak(article['title'])
        speak("The next news is....")
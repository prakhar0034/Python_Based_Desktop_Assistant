import pyttsx3                       # Python Text To Speech
import datetime                      #for date and time 
import speech_recognition as sr      #for recognizing speech from mic
import wikipedia                     #for searching on wikipedia
import webbrowser                    #for searching on web browser
import os                            # operating system
import random                        #randomly choosing
from gnewsclient import gnewsclient  #for google news feed
import requests, json  
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume 
import subprocess                    #for opening pc apps
import screen_brightness_control as sbc
import pyautogui as pi               #for operating mouse and keyboard
import time
import threading

# SETTING VOICES #
engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
# print(voices)
#We have two voices in our system so I have used [0] to use male voice. To see the voices available -- print(voices)
engine.setProperty('voice', voices[1].id) #Setting voice to ([0]male)
NewRate = 190
engine.setProperty('rate', NewRate)

def speak(audio): #created a function (speak) using which NOVA will speak.  
    engine.say(audio)
    engine.runAndWait()


def wishMe():     #NOVA will wish according to the time
    speak('Nova activated')
    hour = int(datetime.datetime.now().hour)
    if hour < 12:
        speak('Good Morning Prakhar!')
    elif hour>=12 and hour<17:
        speak("Good Afternoon Prakhar!")
    else:
        speak("Good Evening Prakhar!")
    

def takeCommand():
    #It takes microphone input from the user and returns string output.
    r = sr.Recognizer()    
    
    with sr.Microphone() as source:
        print('Listening....')
        r.pause_threshold = 1# to view working of such methods press ctrl and click on the desired word.
        r.energy_threshold = 2000
        audio = r.listen(source)
   
    try: #using try becz we don't want our code to crash. If crashes then program will execute on its own
        query = r.recognize_google(audio, language='en-in')
        print('Recognizing....')
        print('Command --> '+query)

    except Exception as e:
        #print(e)It will not look good in the console so removing...
        return "Did not match any results, ask again"
    return query

def Time():
    time = datetime.datetime.now()
    if time.hour < 12:
        meridian = 'am'
    else:
        meridian = 'pm'
    speak(str(time.hour) + str(time.minute) + meridian)

def Weather():
    base_url = "https://api.openweathermap.org/data/2.5/weather?"
    city = 'lucknow'
    api_key = 'c5f5424912f5676d0e64872992f4bb0d'
    full_url = base_url + 'q=' + city + '&appid=' + api_key
    response = requests.get(full_url)
    if response.status_code == 200:
            try:
                data = response.json()
                main = data['main']
                temp = main['temp'] # The temp is floating-point integer (3.1222393339 etc) so convert 
                                    # it to int and then to str fot concatenation.
                humidity = main['humidity'] # in percent but percennt not included; so added
                pressure = main['pressure']
                report = data['weather']
                speak('The temperature is '+ str(int((temp - 273))) + ' degree celsius')
                # speak('The humidity is ' + str(humidity) + 'percent')
                speak("It is "+ report[0]['description'].title() + 'outside')
            except Exception as e:
                print(e)
                speak('Prakhar there was some problem...')

def Date():
    year = int(datetime.datetime.now().year)
    month = int(datetime.datetime.now().month)
    date = int(datetime.datetime.now().day)
    dates_with_st=(1,21,31)
    dates_with_nd=(2,22)
    dates_with_rd=(3,23)
    if date in dates_with_st:
        date = (str(date)+'st')
    elif date in dates_with_nd:
        date = (str(date)+'nd')
    elif date in dates_with_rd:
        date = (str(date)+'rd')
    else:
        date = (str(date)+'th')
    months={1:'January',2:'February',3:'March',4:'April',5:'May',6:'June',7:'July',8:'August',9:'September',10:'October',11:'November',12:'December'}
    for month1 in months:
        if month1 == month:
            speak('It is' + date + months[months1] + str(year))

def VolumeChange(NewRate):
    NewRate = NewRate #Default speaking rate (of pyttsx) is 200. Making it 130.
    engine.setProperty('rate', NewRate) #setting rate to NewRate.

def News():
    if 'business' in query:
        try:
            client = gnewsclient.NewsClient(language='english', location='india', topic='business',max_results=5)
            news_list = client.get_news()
            for item in news_list:
                results2=("Title: ", item['title'])
                speak(results2)
                # print(results2)
                # print('---------------------------------------------------------------------------')
        except Exception as e:
            speak('Prakhar there was some problem..')
            print(e)
    if 'sports' in query:
        try:
            client = gnewsclient.NewsClient(language='english', location='india', topic='sports',max_results=5)
            news_list = client.get_news()
            for item in news_list:
                results2=("Title: ", item['title'])
                speak(results2)
                # print(results2)
                # print('---------------------------------------------------------------------------')
        except Exception as e:
            speak('Prakhar there was some problem..')
            print(e)
    if 'entertainment' in query:
        try:
            client = gnewsclient.NewsClient(language='english', location='india', topic='entertainment',max_results=5)
            news_list = client.get_news()
            for item in news_list:
                results2=("Title: ", item['title'])
                speak(results2)
                # print(results2)
                # print('---------------------------------------------------------------------------')
        except Exception as e:
            speak('Prakhar there was some problem..')
            print(e)
    if 'education' in query:
        try:
            client = gnewsclient.NewsClient(language='english', location='india', topic='education',max_results=5)
            news_list = client.get_news()
            for item in news_list:
                results2=("Title: ", item['title'])
                speak(results2)
                # print(results2)
                # print('---------------------------------------------------------------------------')
        except Exception as e:
            speak('Prakhar there was some problem..')
            print(e)
    if 'health' in query:
        try:
            client = gnewsclient.NewsClient(language='english', location='india', topic='health',max_results=5)
            news_list = client.get_news()
            for item in news_list:
                results2=("Title: ", item['title'])
                speak(results2)
                # print(results2)
                # print('---------------------------------------------------------------------------')
        except Exception as e:
            speak('Prakhar there was some problem..')
            print(e)

def Wikipedia():
    speak("Searching results..")
    query = query.replace("wikipedia", '') #removing wikipedia from query..
    results = wikipedia.summary(query, sentences=2)
    speak('According to wikipedia....')
    print(results)
    speak(results)

def SearchOnInternet():
    try:
        query1 = query.replace('nova','')
        query1 = query.replace('search on internet about','')
        query1 = query1.replace('search on google about')
        query1 = query1.replace('search about')
        query1 = query1.replace('find on google about')
        g_url= 'http://www.google.com/search?q='
        speak('surfing on internet...')
        webbrowser.open(g_url+query1)
    except Exception:
        print(Exception)
        speak('Prakhar there was some problem!')

def PlayMusic():
    music_dir = 'E:\\songs'
    songs = os.listdir(music_dir)
    playing_random = random.choice(songs)
    os.startfile(os.path.join(music_dir, playing_random))

def ChangeBrightness():
    current_brightness = sbc.get_brightness()
    if 'current brightness' in query:
        speak('Current Brightness level is ' + current_brightness)
    elif 'lowest' in query or 'to 0' in query or 'zero' in query:
        speak('Changing brightness to the lowest')
        sbc.set_brightness(0)
    elif 'full' in query or 'brightest' in query or 'highest' in query:
        speak('Changing brightness to highest')
        sbc.set_brightness(100)
    elif '50' in query or 'fifty' in query or 'half' in query:
        speak('Changing brightness..')
        sbc.set_brightness(50)
    elif '10' in query:
        speak('Changing brightness..')
        sbc.set_brightness(10)
    elif '20' in query:
        speak('Changing brightness..')
        sbc.set_brightness(20)
    elif '30' in query:
        speak('Changing brightness..')
        sbc.set_brightness(30)
    elif '40' in query:
        speak('Changing brightness..')
        sbc.set_brightness(40)
    elif '60' in query:
        speak('Changing brightness..')
        sbc.set_brightness(60)
    elif '70' in query:
        speak('Changing brightness..')
        sbc.set_brightness(70)
    elif '80' in query:
        speak('Changing brightness..')
        sbc.set_brightness(80)
    elif '90' in query:
        speak('Changing brightness..')
        sbc.set_brightness(90)

def ChangeSystemVolume():
    # Get default audio device using PyCAW
            devices = AudioUtilities.GetSpeakers()
            interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
            volume = cast(interface, POINTER(IAudioEndpointVolume))
            # Get current volume 
            currentVolumeDb = volume.GetMasterVolumeLevel()
            try:
                if 'decrease' in query or 'down' in query or 'decrease more' in query:
                    volume.SetMasterVolumeLevel(currentVolumeDb - 4.0, None)
                if 'increase' in query or 'up' in query or 'increase more' in query:
                    volume.SetMasterVolumeLevel(currentVolumeDb + 4.0, None)
                if 'full' in query or '100' in query or '100 percent' in query:
                    volume.SetMasterVolumeLevel(currentVolumeDb + 50.0, None)
                if 'zero' in query or '0' in query or '0 percent' in query or 'lowest' in query:
                    volume.SetMasterVolumeLevel(currentVolumeDb - 50.0, None)
                # NOTE: -6.0 dB = half volume !
            except Exception as e:
                print(e)
                speak('Prakhar there was some problem...')

if __name__ == '__main__':
    wishMe()
    while True:
        query = takeCommand().lower()


# GREETINGS STARTING ...
        if 'hello' in query or 'hi' in query:
            speak('Hello Prakhar')
        if 'how are you' in query:
            msg = ['I am good', 'I am fine, thanks', 'I am great']
            speak(random.choice(msg))
            msg1 = ['what about you?', 'how are you Prakhar?']
            speak(random.choice(msg1))
        if 'tell about yourself' in query or 'introduce yourself' in query or 'tell something about yourself' in query:
            speak('Hello everyone, I am an Artificial Intelligence, virtual assisstant, I am capable of doing anything. You can call me Nova 3 point O...Prakhar has created me and I am thankful to him.')
        if 'i am good' in query or 'i am great' in query or 'i am fine' in query or 'same here' in query or 'all good' in query or 'i am great' in query:
            msg2 = ['nice to hear that', 'nice', 'I knew that.......']
            speak(random.choice(msg2))
        if 'who is your god?' in query or 'who is your creator' in query:
            msg3 = ['Prakhar is my creator, he made me', 'Prakhar is my god and creator, he made me']
            speak(random.choice(msg3))
        if 'thank you' in query or 'thanks' in query:
            msg4 = ['your welcome', 'anytime', 'oh no..thats my job']
            speak(random.choice(msg4))
        if 'bored' in query or 'bore' in query:
            speak('Tell me something to do to overcome your boredom')
        if 'cool' in query or 'awesome' in query:
            speak("Thanks")
        if 'new project' in query or 'something new' in query:
            msg5 = ["I hope it's going to be interesting and something new.","Ohh yes, I am interested!"]
            speak(random.choice(msg5))
            msg6 = ['What can I help with?', "What is my work Prakhar?"]
            speak(random.choice(msg6))
        if 'I will let you know' in query or 'instructions' in query:
            speak("OK sure Prakhar")
        if 'ok' in query:
            speak('Hmmmm')


        
# DATE / TIME / WEATHER COMMAND #

        if 'date' in query:
            Date()

        if 'time' in query:
            Time()
            
        if 'weather' in query or 'temperature' in query:
            Weather()


# NEWS COMMANDS STARTING #
        if 'news' in query:
            VolumeChange(130)
            News()
            VolumeChange(190)


# PyAutoGUI Command #
        if 'screenshot' in query:
            pi.hotkey('win', 'prtsc')
        if 'snip' in query:
            pi.hotkey('win', 'shift', 's')
        if 'cut' in query:
            pi.hotkey('ctrl', 'x')
        if 'copy' in query:
            pi.hotkey('ctrl', 'c')
        if 'paste' in query:
            pi.hotkey('ctrl', 'v')
        if 'new tab' in query:
            pi.hotkey('ctrl', 't')
        if 'new window' in query:
            pi.hotkey('ctrl', 'n')
        if 'switch window' in query or 'change window' in query:
            pi.hotkey('alt', 'tab')
        if 'incognito' in query:
            pi.hotkey('ctrl', 'shift', 'n')
        if 'switch tab' in query or 'change tab' in query:
            pi.hotkey('ctrl', 'tab')
        if 'notification' in query:
            pi.click(1878, 1060)
        if 'wifi' in query or 'WiFi' in query or 'Wi-Fi' in query:
            pi.click(1722, 1054)
        if 'battery' in query:
            pi.click(1693, 1059)
        if 'sound' in query or 'speaker' in query:
            pi.click(1750, 1058)
        if 'reload' in query or 'refresh' in query:
            pi.hotkey('ctrl', 'r')
        if 'google history' in query or 'history' in query:
            pi.hotkey('ctrl', 'h')
        if 'downloads' in query:
            pi.hotkey('ctrl', 'j')
        if 'desktop' in query:
            pi.hotkey('win', 'd')
        
        
# WIKI AND WEB BROWSER SEARCHES #
        if 'wikipedia' in query:
            wikipedia()

        if 'search on internet' in query or 'search on google' in query or 'find on google' in query or 'search about' in query:
            SearchOnInternet()
           

# NECESSARY WEBSITES THROUGH WEB #
        if 'open google' in query.lower():
            webbrowser.open('//google.com')
            speak('Opening google...')

        if 'open youtube' in query.lower():
            webbrowser.open("http://youtube.com")
            speak('Opening youtube...')

        if 'open whatsapp web' in query.lower():
            webbrowser.open('http://web.whatsapp.com')
            speak('Opening whatsapp web...')

        if 'open google classroom' in query.lower():
            webbrowser.open('https://classroom.google.com/u/1/h')
            speak('Opening Google Classroom...')

        if 'open gmail' in query.lower():
            webbrowser.open('http://mail.google.com')
            speak('Opening Gmail...')


# MUSIC & VOLUME #

        if 'play music' in query or 'next music' in query or 'some music' in query: #return random.choice(GREETING_RESPONSES)+'.'
            PlayMusic()

        if 'volume' in query:
            ChangeSystemVolume()


# OPENING PC APPS #

        if 'calculator' in query:
            subprocess.Popen('C:\\Windows\\System32\\calc.exe')
            speak('Opening Calculator')

        if 'microsoft word' in query or 'open word' in query:
            subprocess.Popen('C:\Program Files (x86)\Microsoft Office\Office14\WINWORD.exe')
            speak('Opening Office')

        if 'microsoft excel' in query or 'excel' in query:
            subprocess.Popen('C:\Program Files (x86)\Microsoft Office\Office14\EXCEL.exe')
            speak('Opening Excel')
        

# CLASSROOMS / STUDY RELATED STUFF #        

        if 'maths solution' in query or 'ml agarwal solution' in query:
            webbrowser.open('https://www.aplustopper.com/ml-aggarwal-icse-solutions-for-class-9-maths/')
            speak('opening Maths Solutions...')
       
        if 'open physics class' in query:
            webbrowser.open('https://classroom.google.com/u/1/c/MTQxMTIzODY0MDE0')
            speak('opening physics class..')

        if 'open history class' in query:
            webbrowser.open('https://classroom.google.com/u/1/c/MTEwMzYxNzI1MjIz')
            speak('opening history class...')

        if 'open english class' in query:
            webbrowser.open('https://classroom.google.com/u/1/c/MTAxMDIwMzA5MDMw')
            speak('opening english class..')

        if 'open maths class' in query:
            webbrowser.open('https://classroom.google.com/u/1/c/NzEyNTE0NzE4NTFa')
            speak('opening maths class...')

        if 'open chemistry class' in query:
            webbrowser.open('https://classroom.google.com/u/1/c/MTIyMDE1MjY2MjYw')
            speak('opening chemistry class..')

        if 'open biology class' in query:
            webbrowser.open('https://classroom.google.com/u/1/c/MTAwODEyODc2NzAz')
            speak('opening biology class..')

        if 'open computer class' in query:
            webbrowser.open('https://classroom.google.com/u/1/c/OTU2NTUwODE5ODda')
            speak('opening computer class..')

        if 'open geography class' in query:
            webbrowser.open('https://classroom.google.com/u/1/c/MTI1NzQ4MjU4MTE0')
            speak('opening geography class..')

        if 'open hindi class' in query:
            webbrowser.open('https://classroom.google.com/u/1/c/NTYxMjU2MjQyNDBa')
            speak('opening hindi class..')


# BRIGHTNESS CONTROL #
        if 'brightness' in query:
            ChangeBrightness()

# OS COMMANDS (ShutDown) #
        if 'shutdown the laptop' in query or 'shutdown the computer' in query or 'shutdown the pc' in query:
            speak('Shutting down in 8 seconds')
            os.system('shutdown /s /t 8')

# TERMINATING NOVA #
        if 'exit' in query or 'bye nova' in query or 'end working' in query:
            speak("nova terminating")
            quit()


# NOVA ENDED #
import pyttsx3 #pip install pyttsx3
import speech_recognition as sr  #pip install SpeechRecognition
import datetime #pip install DateTime
import wikipedia #pip install wikipedia
import webbrowser #pip install webbrowser
import pywhatkit  #pip install pywhatkit
import pyjokes  #pip install pyjokes
import os       #pip install os-sys
import pyautogui   #pip install PyAutoGUI
import subprocess  #pip install subprocess.run
import winshell  #pip install winshell
import ctypes    #pip install ctypes-callable
  



engine = pyttsx3.init('sapi5') #init function to get an engine instance for the speech synthesis
#sapi5 - SAPI5 on Windows
#nsss - NSSpeechSynthesizer on Mac OS X
#espeak - eSpeak on every other platform  
voices = engine.getProperty('voices')

engine.setProperty('voice', voices[1].id)


def speak(audio):
    engine.say(audio) # say method on the engine that passing input text to be spoken
    engine.runAndWait() # run and wait method, it processes the voice commands.


def wishMe():
    hour = int(datetime.datetime.now().hour)
    if hour>=0 and hour<12:
        speak("Good Morning!")

    elif hour>=12 and hour<18:
        speak("Good Afternoon!")   

    else:
        speak("Good Evening!")  

    speak("Abdullah I am your Window Assistant. Please tell me how may I help you")       

def takeCommand():
    #It takes microphone input from the user and returns string output

    r = sr.Recognizer()
    with sr.Microphone() as source: #use the microphone as source for input. 
    
        r.adjust_for_ambient_noise(source, duration=1)  #wait for a second to let the recognizer adjust the 
        
        r.dynamic_energy_threshold = True     #energy threshold based on the surrounding noise level
        print("Listening...")  
        audio = r.listen(source)  #listens for the user's input

    try:
        print("Recognizing...")    
        query = r.recognize_google(audio, language='en-in')
        print(f"User said: {query}\n")

    except Exception as e:   #error occurs when google could not understand what was said
        # print(e)    
        print("Abdullah Say that again please...")  
        return "None"
    return query



if __name__ == "__main__":
    wishMe()
    while True:
    # if 1:
        query = takeCommand().lower()

        # Logic for executing tasks based on query
        if 'wikipedia' in query:
            speak('Searching Wikipedia...')
            query = query.replace("wikipedia", "")
            results = wikipedia.summary(query, sentences=2)
            speak("According to Wikipedia")
            print(results)
            speak(results)


        elif 'open youtube' in query:
            speak('Abdullah I am opening youtube')
            webbrowser.open("https://youtube.com")


        elif 'open google' in query:
            speak('Abdullah I am opening google')
            webbrowser.open("https://google.com")


        elif 'open lms' in query:
            speak('Abdullah I am lms')
            webbrowser.open("http://lms.mnsuam.edu.pk/")


        elif 'check email' in query:
            speak('Abdullah I am checking email')
            webbrowser.open("https://mail.google.com/mail/u/0/?tab=rm&ogbl#inbox")


        elif 'open stackoverflow' in query:
            speak('Abdullah I am opening stackoverflow')
            webbrowser.open("https://stackoverflow.com") 



        elif 'play' in query:
            song=query.replace('play','')
            speak('Abdullah i am playing'+ song)
            pywhatkit.playonyt(song)


        elif 'joke' in query:
             joke=pyjokes.get_joke()
             print(joke) 
             speak(joke)  
             
           

        elif 'what is the time' in query:
            strTime = datetime.datetime.now().strftime("%H:%M:%S")    
            speak(f"Abdullah, the time is {strTime}")

     

        elif 'search' in query:
            show = query.replace('search', '')
            speak('Abdullah I am searching' + show)
            url = 'https://google.com/search?q=' + show
            webbrowser.get().open(url)   


        elif 'music' in query:
          play_music = query.replace('music', '')
          speak('Abdullah I am playing music')
          music_dir = "E:\\music\\"
          songs = os.listdir(music_dir)
          print(songs)
          os.startfile(os.path.join(music_dir,songs[0]))  



        elif 'location' in query:
         location = query.replace('location', '')
         speak('Abdullah I am locating' + location)
         url = 'https://google.nl/maps/place/' + location + '/&map;'
         webbrowser.get().open(url)



        elif "where is" in query:
            query = query.replace("where is", "")
            location = query
            speak("Abdullah I am locating")
            speak(location)
            webbrowser.open("https://www.google.nl/maps/place/" + location + "")


        elif 'open code' in query:
            speak('Abdullah opening visual studio code')
            codePath = "C:\\Users\\samee\\AppData\\Local\\Programs\\Microsoft VS Code\\Code.exe"
            os.startfile(codePath)


        elif 'open word' in query:
            speak('Abdullah opening Microsoft Word')
            codePath = "C:\\Program Files\\Microsoft Office\\root\\Office16\\WINWORD.EXE"
            os.startfile(codePath) 

        elif 'excel' in query:
            speak('Abdullah opening Microsoft Excel')
            codePath = "C:\\Program Files\\Microsoft Office\\root\\Office16\\EXCEL.EXE"
            os.startfile(codePath)

        elif 'powerpoint' in query:
            speak('Abdullah opening Microsoft PowerPoint')
            codePath = "C:\\Program Files\\Microsoft Office\\root\\Office16\\POWERPNT.EXE"
            os.startfile(codePath)


        elif 'projects' in query:
            speak('Abdullah opening Projects folder')
            codePath = "I:\\Python Project"
            os.startfile(codePath)

        elif 'drive i' in query:
            speak('Abdullah opening drive i')
            codePath = "I:"
            os.startfile(codePath)   

        elif 'drive d' in query:
            speak('Abdullah opening drive d')
            codePath = "D:"
            os.startfile(codePath)    

        elif 'drive e' in query:
            speak('Abdullah opening drive e')
            codePath = "e:"
            os.startfile(codePath)   
                  

        elif 'screenshot' in query:
            image = pyautogui.screenshot()
            image.save('screenshot.png')
            speak('Screenshot taken.')         


    


        elif "log off" in query or "sign out" in query:
            speak("Ok , your pc will log off in 10 sec make sure you exit from all applications")
            subprocess.call(["shutdown", "/l"])
  

        elif 'change background' in query:
          ctypes.windll.user32.SystemParametersInfoW(20,0,"C:\\Users\\samee\\OneDrive\\Pictures\\Screenshots",0)
          speak("Background changed successfully")



        elif 'lock window' in query:
                speak("locking the device")
                ctypes.windll.user32.LockWorkStation()


        elif 'shutdown system' in query:
                speak("Hold On a Sec ! Your system is on its way to shut down")
                subprocess.call('shutdown / p /f')


        elif 'empty recycle bin' in query:
            winshell.recycle_bin().empty(confirm = False, show_progress = False, sound = True)
            speak("Recycle Bin Recycled")


        elif "restart" in query:
            subprocess.call(["shutdown", "/r"]) 


        elif "hibernate" in query or "sleep" in query:
            speak("Hibernating")
            subprocess.call("shutdown / h")      



        elif 'exit' in query:
         b = 'Abdullah Good Bye'
         speak(b)
         print(b)
         exit()   

        else:
           a = 'Abdullah I do not understand please say it again '
           speak(a)
           print(a) 
            



        




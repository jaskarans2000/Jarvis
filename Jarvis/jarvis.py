import pyttsx3
import datetime
import speech_recognition as sr
import wikipedia
import webbrowser
import os
import random
import smtplib
import xlrd
import pyautogui 
import psutil
import pyjokes


#Initializes the engine and sapi5 let us choose microsoft speech api
engine = pyttsx3.init('sapi5')
#voices contains the list of available voices on your pc
voices=engine.getProperty('voices')
#print(voices[0].id)
#this sets the voice to male or female depending on index 0 or 1
engine.setProperty('voice', voices[1].id)

def getonline():
    speak("Connecting to remote servers")
    speak("Establishing seure connections")
    speak('Starting all systems')
    speak('Downloading and installing all required drivers')
    speak('Just a second sir')
    speak(' ')
    os.startfile("C:\\Program Files\\Rainmeter\\Rainmeter.exe")
    speak(' ')
    speak('Secure connection established')
    speak('All systems have started')
    speak('All drivers are securely installed on your private servers')
    speak('Now I am online sir')

def speak(audio):
    '''This function takes a string input and gives us voice output or 
    basically speaks whatever we want it to speak
    '''
    engine.say(audio)
    engine.runAndWait()

def wishme():
    '''This function checks the hour of day and greets accordingly.
    '''
    hour=int(datetime.datetime.now().hour)
    if hour>=4 and hour<12:
        speak("Good Morning!")

    elif hour>=12 and hour<18:
        speak("Good Afternoon!")

    elif hour>=18 and hour<20:
        speak("Good Evening!")

    else :
        speak("Good Night!")

    speak("I am Jarvis Sir.")
    Time()
    Date()
    speak("How can I help you?")

def takecommand():
    '''This function takes command as voice from user microphone and 
    returns string output
    '''
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Say something!")
        #speak("Listening Sir.")
        r.pause_threshold=1
        r.energy_threshold=400
        audio = r.listen(source)
    try:
        print("Recognizing...")
        query=r.recognize_google(audio, language='en-in')
        print(f"User said : {query}\n")

    except Exception as e:
        print(e)
        print("\n Say that again please...")
        return "None"
    return query

def Wikipedia():
    speak("Searching Wikipedia...")
    que=query.replace("wikipedia","")
    result=wikipedia.summary(que,sentences=2)
    speak("According to Wikipedia...")
    speak(result)

def Youtube():
    speak("Opening Youtube...")
    webbrowser.open("youtube.com")

def Google():
    speak("Opening Google...")
    webbrowser.open("google.com")

def StackOverflow():
    speak("Opening Stack Overflow...")
    webbrowser.open("stackoverflow.com")

def Music():
    music_dir='F:\\Music\\Favourite Music'
    songs=os.listdir(music_dir)
    print(songs)
    os.startfile(os.path.join(music_dir,random.choice(songs)))

def Time():
    sTrtime=datetime.datetime.now().strftime("%H:%M:%S")
    speak(f"Sir, the time is {sTrtime}.")

def VSCode():
    codePath = "C:\\Users\\lenovo\\AppData\\Local\\Programs\\Microsoft VS Code\\Code.exe"
    speak("Opening V S Code")
    os.startfile(codePath)

def AndroidStudio():
    codePath = "C:\\Program Files\\Android\\Android Studio1\\bin\\studio64.exe"
    os.startfile(codePath)

def IntelliJ():
    codePath = "C:\\Program Files\\JetBrains\\IntelliJ IDEA Community Edition 2018.3.4\\bin\\idea64.exe"
    os.startfile(codePath)

def sendEmail(to,content):
    wb = xlrd.open_workbook("contacts.xlsx") 
    sheet = wb.sheet_by_index(0) 
    print(sheet.cell_value(1, 1))
    file1 = open("emaildata.txt","r+")
    lines=file1.readlines()
    file1.close()  
    #print(to)
    mail=""
    for i in range(1,sheet.nrows):
        if sheet.cell_value(i,0).lower()==to:
            mail=sheet.cell_value(i,1)
            print(mail)
            break 
    server=smtplib.SMTP('smtp.gmail.com',587)
    server.ehlo()
    server.starttls()
    server.login(lines[0],lines[1])
    server.sendmail(lines[0],mail,content)
    server.close()

def Date():
    a=datetime.datetime.now()
    x=str(a)
    speak("Sir, today is")
    speak(a.strftime("%A"))
    y=x.split(" ")
    speak(y[0])

def Search():
    speak("What you want to search?")
    content=takecommand().lower()
    chromePath=":\\Program Files (x86)\\Google\\Chrome\\chrome.exe %s"
    webbrowser.get(chromePath).open_new_tab(content+".com")

def Shut():
    os.system("shutdown /s /t 1")

def Restart():
    os.system("shutdown /r /t 1")

def Logout():
    os.system("shutdown -l")

def Screenshot():
    speak("Taking Screenshot...")
    img=pyautogui.screenshot()
    img.save(r'C:\\Users\\lenovo\\Desktop\\screenshot1.png')
    speak("Screenshot taken...")

def cpu():
    usage= str(psutil.cpu_percent())
    speak("CPU is at"+usage+"percent")
    batter=psutil.sensors_battery()
    speak("Battery is at:")
    speak(batter.percent)
    speak("percent")

def jokes():
    speak(pyjokes.get_joke())

if __name__ == "__main__":
    #os.startfile("jarvisimg.jpg")
    getonline()
    wishme()
    while True:

        query=takecommand().lower()

        #Logic for executing task based on query

        if 'wikipedia' in query:
            Wikipedia()

        elif 'open youtube' in query:
            Youtube()

        elif 'open google' in query:
            Google()

        elif 'open stack overflow' in query:
            StackOverflow()

        elif 'play music' in query:
            Music()

        elif 'time' in query:
            Time()                
        
        elif 'date' in query:
            Date()

        elif 'open code' in query:
            VSCode()     

        elif 'open android studio' in query:
            AndroidStudio()     

        elif 'open idea' in query:
            IntelliJ()               

        elif 'email' in query:
            try:
                speak("To whom you want to send email")
                to=takecommand().lower()
                speak("What should I say?") 
                content=takecommand()
                sendEmail(to,content)
                speak("Email has been sent")
            except Exception as e:
                print(e)
                speak("Sorry Sir, some problem has occured and I am unable to send email currently!")

        elif 'offline' in query:
            exit()

        elif 'search' in query:
            Search()
            
        elif 'shut down' in query:
            Shut()

        elif 'restart' in query:
            Restart()

        elif 'logout' in query:
            Logout()

        elif 'screenshot' in query:
            Screenshot()

        elif 'cpu' in query:
            cpu()

        elif 'joke' in query:
            jokes()

        elif 'how are you'in query or 'what \'s up' in query:
            lis=['I am cool, what about you?','Just doing my work','Performing my duty of serving you','I am nice and full of energy']
            speak(random.choice(lis))


    #speak("Jaskaran is a good boy") 
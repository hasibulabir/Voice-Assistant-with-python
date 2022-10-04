import email
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from fileinput import filename
from tkinter import PAGES
#from flask import request
#import instaloader
import pyttsx3
import speech_recognition as sr
import datetime
import os
import cv2
import random
from requests import get
import webbrowser
import wikipedia
import pywhatkit as kit
import smtplib
import sys
import time
import pyjokes
import PyQt5
import pyautogui
import PyPDF2







engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voices',voices[0].id)

def speak(audio):
    engine.say(audio)
    print(audio)
    engine.runAndWait()

def takecommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("listening...")
        r.pause_threshold = 1
        audio = r.listen(source,timeout=1,phrase_time_limit=5)

    try :
        print("Recognozing...")
        query = r.recognize_google(audio,language='en-in')
        print(f"user said : {query}")

    except Exception as e :
        speak("Say that again please...")
        return "none"
    return query

def wish():
    hour = int(datetime.datetime.now().hour)
    tt= time.strftime("%I:%M %p")

    if hour >= 0 and hour <= 12 :
        speak(f"Good Morning, its {tt}")
    elif hour > 12 and hour < 18 :
        speak(f"Good Afternoon, its {tt}")
    else :
        speak(f"Good Evening, its {tt}")
    speak("I am your voice assistant sir. Please tell me how can i help you")

    
#for news updates
    # importing requests package
import requests	

def news():
	
	# BBC news api
	# following query parameters are used
	# source, sortBy and apiKey
	query_params = {
	"source": "bbc-news",
	"sortBy": "top",
	"apiKey": "4dbc17e007ab436fb66416009dfb59a8"
	}
	main_url = " https://newsapi.org/v1/articles"

	# fetching data in json format
	res = requests.get(main_url, params=query_params)
	open_bbc_page = res.json()

	# getting all articles in a string article
	article = open_bbc_page["articles"]

	# empty list which will
	# contain all trending news
	results = []
	
	for ar in article:
		results.append(ar["title"])
		
	for i in range(len(results)):
		
		# printing all trending news
		print(i + 1, results[i])

	#to read the news out loud for us
	from win32com.client import Dispatch
	speak = Dispatch("SAPI.Spvoice")
	speak.Speak(results)	
 
 # importing the required modules

def pdf_reader():
    book = open('Kevin knight.pdf','rb')
    pdfReader = PyPDF2.pdfFileReader(book)
    pages =pdfReader.numPages
    speak(f"Total numbeers of pages in this book{pages}")
    speak("sir please enter the page number i have to read")
    pg = int(input("please enter the page number: "))
    page = pdfReader.getPage(pg)
    text = page.extractText()




if __name__ == "__main__" :

    wish()
    while True:
    #if 1 :
        query = takecommand().lower()

        if "open notepad" in  query:
            npath="C:\\Windows\\system32\\notepad"
            os.startfile(npath)

        elif "open adobereader" in  query:
            apath="C:\\Program Files (x86)\\Adobe\\Reader 11.0\\Reader\\AcroRd32.exe"
            os.startfile(apath)

        elif "open command prompt" in query:
            os.system("start cmd")

        elif "open camera" in query:
            cap = cv2.VideoCapture(0)
            while True :
                ret , img= cap.read()
                cv2.imshow('webcam',img)
                k = cv2.waitKey(50)
                if k == 27 :
                    break
            cap.release()
            cv2.destroyAllWindows()

        elif "open music" in query:
            music_dir = ""
            songs = os.listdir(music_dir)
            for song in songs :
                if song.endswith('.mp3') :
                    os.startfile(os.path.join(music_dir,song))
        
        elif "ip address" in query:
            ip=get('https://api.ipify.org')
            speak(f"your ip address is {ip}")

        elif "wikipedia" in query :
            speak("searching wikipedia...")
            query = query.replace("wikipedia","")
            results = wikipedia.summary(query, sentences=2)
            speak("according to wikipedia")
            speak(results) 

        elif "open youtube" in query :
            webbrowser.open("www.youtube.com")

        elif "open facebook" in query :
            webbrowser.open("www.facebook.com")

        elif "open stackoverflow" in query :
            webbrowser.open("www.stackoverflow.com")

        elif "open google" in query :
            speak("Sir, What should i search on google")
            cm = takecommand().lower()
            webbrowser.open(f"{cm}")

        elif "send whats message" in query :
            kit.sendwhatmsg("+8801521418384", "this is testing protocol",4,13)
            time.sleep(120)
            speak("massage has been sent")

      
        
        elif "you can sleep" in query :
            speak("thanks for using me sir, have a good day.")
            sys.exit()
            
        elif "close notepad" in query :
            speak("okey sir, closing notepad")
            os.system("taskkill /f /im notepad.exe")
            
        elif "set alarm" in query :
            nn = int(datetime.datetime.now().hour)
            if nn==22:
                music_dir=""
                songs= os.listdir(music_dir)
                os.startfile(os.path.join(music_dir, songs[0]))
        
        elif "tell me a joke" in query :
            joke = pyjokes.get_joke()
            speak(joke)
        
        elif "shut  down the system" in query :
            os.system("shutdown /s /t 5")
            
        elif "restart the system" in query :
            os.system("shutdown /r /t 5")
            
        elif "sleep the system" in query :
            os.system("rundll32.exe powrprof.dll,SetSuspendState 0,1,0")
            
            #video part-3
        elif "switch the window" in query:
            pyautogui.keyDown("alt")
            pyautogui.press("tab")
            time.sleep(1)
            pyautogui.keyUp("alt")
        elif "tell me news" in query:
            speak("please wait sir,fetching the latest news")
            news()              
        elif "send email" in query:
             speak("what should i say?")
             query = takecommand().lower()
             if "send a file" in query:
                 email = 'hasibul.abir4@gmail.com' #your email
                 password = 'your pass' #your email account password
                 send_to_email = 'sarowerhasan.cse@gmail.com' #whom you are sending the message to
                 speak("okay sir, what is the subject for this email")
                 query = takecommand().lower()
                 subject = query # the subjectg in the email
                 speak("and sir, what is the message3 for this email")
                 query2 = takecommand().lower()
                 message = query2 
                 speak("sir please enter the correct path of the file in the shell")
                 file_location = input("please enter the path here")
                 
                 speak("please wait, i am sending email now")
                 
                 msg = MIMEMultipart()
                 msg['From']= email
                 msg['To'] = send_to_email
                 msg['Subject'] = subject
                 
                 
                 msg.attach(MIMEText(message,'plain'))  #import mimetext????????
                 
                 #setup the attachment
                 
                 filename = os.path.basename(file_location)
                 attachment = open(file_location,"rb")
                 part = MIMEBase('application','octet-stream')
                 part.set_payload(attachment.read())
                 encoders.encode_base64(part)
                 part.add_header('content_Disposition', "attachment; filename = %s"%filename)
                 
                 #attach the attachment to the MIMEMultipart object
                 msg.attach(part)
                 
                 server = smtplib.SMTP('smtp.gmail.com',587)
                 server.starttls()
                 server.login(email,password)
                 text =msg.as_string()
                 server.sendmail(email, send_to_email,text)
                 server.quit()
                 speak("Email has been sent ")
                 
                 
             else:
                 email = 'your@gmail.com'
                 password = 'yuour_pass'
                 send_to_email ='person@gmail.com'
                 message = query
                 
                 
                 server = smtplib.SMTP('smtp.gmaiil.com',587)
                 server.starttls()
                 server.login(email,password)
                 server.sendmail(email, send_to_email,message)
                 server.quit()
                 speak("Email has been sent to abir")
        #Jarvis Video part 5
        elif "where i am" in query or "where we are " in query:
            speak("wait sir, let me check")
            try:
                ipAdd = requests.get('https://api.ipify.org').text
                print(ipAdd)
                url = 'https://get.geojs.io/v1/ip/geo'+ipAdd+'.json'
                geo_requests = request.get(url)
                geo_data = geo_requests.json()
                
                city = geo_data['city']
                country = geo_data['country']
                speak(f"sir i am not sure, but i think we are in {city} city of {country} country")
            except Exception as e:
                speak("sorry sir, Due to network issue i am not able to find where we are.")
                pass
            
        
            
            
        elif "take screenshot" in query or "take a screenshot" in query:
            speak("sir, pleasek tell me the name for this screenshot file")
            name = takecommand().lower()
            speak("Please sir hold the screen for few seconds, i am taking screenshot")
            time.sleep(3)
            img = pyautogui.screenshot()
            img.save(f"{name}.png")
            speak("i am done sir, the screenshot is saved in out main folder.now i am ready for the next comand")
            speak("sir , do you have any other work")
            
            
            #Jarvis video-6
        elif "read pdf" in query:
            pdf_reader()
            
        #jarvis video 7
        elif "hide all files" in query or "hide this folder" in query or "visible " in query:
            speak("sir please tell me you want to hide this folder or make it visible for everyone")
            condition = takecommand().lower()
            if "hide" in condition:
                os.system("attrib +h /s /d")
                speak("sir, all the files in this folder are now hidden")
                
                
            elif "show" in condition:
                os.system("attrib -h /s /d")
                speak("sir, all the files in this folder are now visible to everyone.I wish you are taking")
                
                
            elif "leave it" in condition or  "leave for now " in condition:
                speak("Ok sir")
                
            
            
            
                 
            
            
             
            
            
            
            

        
                

















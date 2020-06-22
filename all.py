from __future__ import print_function
import json
import re
import datetime
import pickle
import os.path
import requests
import pywhatkit as kit
import pytz
import speech_recognition as sr  # importing speech recognition package from google api
import playsound    # to play saved mp3 file
import os   # to save/open files
import wolframalpha # to calculate strings into formula, its a website which provides api, 100 times per day
import pyttsx3
import googletrans
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

from selenium import webdriver  # to control browser operations
from selenium.webdriver.common.keys import Keys
from io import BytesIO
from io import StringIO
from time import ctime
from textblob import TextBlob
from gtts import gTTS
from googletrans import Translator

num = 1
SCOPES = ['https://www.googleapis.com/auth/calendar.readonly']
MONTHS = ['january', 'february', 'march', 'april', 'may', 'june', 'july', 'august', 'september', 'october', 'november', 'december']
DAYS = ['monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday', 'sunday']
DAY_EXTENSIONS = ["nd", "rd", "th", "st"]
res = requests.get('https://ipinfo.io/')
data = res.json()

city = data['city']
location = data['loc'].split(',')
latitude = location[0]
longitude = location[1]

API_Key = "t8UHCPseDMoc"
Project_Token = "tjcZ0CVvV5jd"
Run_Token = "tDL3XAvPnETT"


def assistant_speaks(output):
	print("Reply: ", output)
	toSpeak = pyttsx3.init()
	toSpeak.say(output)
	toSpeak.runAndWait()


def get_audio():
    r = sr.Recognizer()
    audio = ''
    with sr.Microphone() as source:
        print("Speak...")
        audio = r.listen(source, phrase_time_limit=5)
    print("Stop.")
    try:
        text = r.recognize_google(audio,language='en-In')
        print("You : ", text)
        return text.lower()
    except:
        assistant_speaks("Could not understand your audio, PLease try again!")
        return 0


def search_web(input):
    driver = webdriver.Chrome()
    driver.implicitly_wait(1)
    driver.maximize_window()
    try:
	    
	    if 'google' in input:
	    	indx = input.lower().split().index('google')
	    	query = input.split()[indx + 1:]
	    	driver.get("https://www.google.com/search?q=" + '+'.join(query))
	    elif 'search' in input:
	        indx = input.lower().split().index('google')
	        query = input.split()[indx + 1:]
	        driver.get("https://www.google.com/search?q=" + '+'.join(query))
	    else:
	    	driver.get("https://www.google.com/search?q=" + '+'.join(input.split()))
	    	return
    except Exception as e:
        print(e)


def open_application(input):
    if "chrome" in input:
        assistant_speaks("Google Chrome")
        os.startfile('C:\Program Files (x86)\Google\Chrome\Application\chrome.exe')
        return
    elif "firefox" in input or "mozilla" in input:
        assistant_speaks("Opening Mozilla Firefox")
        os.startfile('C:\Program Files\Mozilla Firefox\\firefox.exe')
        return
    elif "word" in input:
        assistant_speaks("Opening Microsoft Word")
        os.startfile('C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Office 2013\\Word 2013.lnk')
        return
    elif "excel" in input:
        assistant_speaks("Opening Microsoft Excel")
        os.startfile('C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Office 2013\\Excel 2013.lnk')
        return
    else:
        assistant_speaks("Application not available")
        return


def process_text(input):
	try:
	    if "who are you" in input or "define yourself" in input:
	        speak = '''Hello, I am Person. Your personal Assistant.
	        I am here to make your life easier. 
	        You can command me to perform various tasks such as calculating sums or opening applications etcetra'''
	        assistant_speaks(speak)
	        return
	    elif "who made you" in input or "created you" in input:
	        speak = "I have been created by Sayan Bhattacharjee."
	        assistant_speaks(speak)
	        return
	    elif "crazy" in input:
	        speak = """Well, there are 2 mental asylums in India."""
	        assistant_speaks(speak)
	        return
	    elif "calculate" in input.lower():
	        app_id= "E46YXW-T5LG6RT7K7"
	        client = wolframalpha.Client(app_id)

	        indx = input.lower().split().index('calculate')
	        query = input.split()[indx + 1:]
	        res = client.query(' '.join(query))
	        answer = next(res.results).text
	        assistant_speaks("The answer is " + answer)
	        return
	    elif 'youtube' in input or 'play'  in input:
	    	indx = input.lower().split()
	    	query = indx[1:]
	    	kit.playonyt("".join(query)) 
	    	return   
	    elif 'open' in input:
	        open_application(input.lower())
	        return
	    elif 'search' in input:
	        search_web(input.lower())
	        return
	    elif 'location' in input:
	    	assistant_speaks("Your location is :" + city)
	    	return
	    elif 'what do i have' in input or 'do i have plans' in input or 'am i busy' in input:
	    	date = get_date(text)
	    	if date:
	    		get_events(date, service)
	    	else:
	    		assistant_speaks("I didn't quite get that")
	    	return
	    elif 'forecast' in input :
	    	 driver.get("https://www.weather-forecast.com/locations/"+city+"/forecasts/latest")
	    	 assistant_speaks(driver.find_elements_by_class_name("b-forecast__table-description-content")[0].text)
	    	 return
	    elif 'weather' in input:
        	indx = input.lower().split()
        	word = indx[-1]
        	driver.get("https://www.weather-forecast.com/locations/"+word+"/forecasts/latest")
        	assistant_speaks(driver.find_elements_by_class_name("b-forecast__table-description-content")[0].text)
        	return

	    elif 'virus' in input:
	    	assistant_speaks('What information you want about corona virus ?')
	    	text = get_audio()
	    	corona_virus(text)
	    elif 'translate' in input:
	    	fun_translate()

	    else:
	    	assistant_speaks("I can search the web for you, Do you want to continue?")
	    	ans = get_audio()
	    	if 'yes' in str(ans) or 'yeah' in str(ans):
	    		search_web(input)
	    	else:
	    		return
	except Exception as e:
		print(e)
		assistant_speaks("I don't understand, I can search the web for you, Do you want to continue?")
		ans = get_audio()
		if 'yes' in str(ans) or 'yeah' in str(ans):
			search_web(input)


def Authenticate_google():
    
    creds = None
    
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('calendar', 'v3', credentials=creds)
    return service

def get_events(day, service):
    date = datetime.datetime.combine(day, datetime.datetime.min.time())
    end_date = datetime.datetime.combine(day, datetime.datetime.max.time())
    utc = pytz.UTC
    date = date.astimezone(utc)
    end_date = end_date.astimezone(utc)

    events_result = service.events().list(calendarId='primary', timeMin=date.isoformat(), timeMax=end_date.isoformat(),
                                        singleEvents=True,
                                        orderBy='startTime').execute()
    events = events_result.get('items', [])

    if not events:
        assistant_speaks('No upcoming events found.')
    else:
        assistant_speaks(f"You have {len(events)} events on this day.")

        for event in events:
            start = event['start'].get('dateTime', event['start'].get('date'))
            print(start, event['summary'])
            start_time = str(start.split("T")[1].split("-")[0])
            if int(start_time.split(":")[0]) < 12:
                start_time = start_time + "am"
            else:
                start_time = str(int(start_time.split(":")[0])-12)
                start_time = start_time + "pm"

            assistant_speaks(event["summary"] + " at " + start_time)
def get_date(text):
	text = text.lower()
	today = datetime.date.today()
	if text.count("today") > 0:
		return today

	day = -1
	day_of_week = -1
	month = -1
	year = today.year

	for word in text.split():
		if word in MONTHS:
			month = MONTHS.index(word) + 1
		elif word in DAYS:
			day_of_week = DAYS.index(word)
		elif word. isdigit():
			day = int(word)
		else:
			for ext in DAY_EXTENSIONS:
				found = word.find(ext)
				if found > 0:
					try:
						day = int(word[:found])
					except:
						pass
		if month < today.month and month != -1:
			year = year+1

		if month == -1 and day != -1:
			if day < today.day:
				month = today.month + 1
			else:
				month = today.month

		if month == -1 and day == -1 and day_of_week != -1:
			current_day_of_week = today.weekday()
			dif = day_of_week - current_day_of_week

			if dif < 0:
				dif += 7
				if text.count("next") >= 1:
					dif += 7

			return today + datetime.timedelta(dif)

		if day != -1:
			return datetime.date(month=month, day=day, year=year)
service = Authenticate_google()

class Data:
	def __init__(self, api_key, project_token):
		self.api_key = api_key
		self.project_token = project_token
		self.params = {
			"api_key": self.api_key
		}
		self.get_data()

	def get_data(self):
		response = requests.get(f'https://www.parsehub.com/api/v2/projects/{Project_Token}/last_ready_run/data', params={"api_key": API_Key})
		self.data = json.loads(response.text)

	def get_total_cases(self):
		data = self.data['total']

		for content in data:
			if content['name'] == "Coronavirus Cases:":
				return content['value']
	
	def get_total_deaths(self):
		data = self.data['total']

		for content in data:
			if content['name'] == "Deaths:":
				return content['value']

	def get_total_recovered(self):
		data = self.data['total']

		for content in data:
			if content['name'] == "Recovered:":
				return content['value']

	def get_country_data(self, country):
		data = self.data['country']

		for content in data:
			if content['name'].lower() == country.lower():
				return content
		return "0"

	def get_list_of_countries(self):
		countries = []
		for country in self.data['country']:
			countries.append(country['name'].lower())
		return countries

def corona_virus(text):
	data = Data(API_Key, Project_Token)
	country_list = data.get_list_of_countries()

	TOTAL_PATTERNS = {
					re.compile("[\w\s]+ total [\w\s] + cases"):data.get_total_cases,
					re.compile("[\w\s]+ total cases"): data.get_total_cases,
                    re.compile("[\w\s]+ total [\w\s]+ deaths"): data.get_total_deaths,
                    re.compile("[\w\s]+ total deaths"): data.get_total_deaths,
                    re.compile("[\w\s]+ total [\w\s]+ recovered"): data.get_total_recovered,
                    re.compile("[\w\s]+ total recovered"): data.get_total_recovered,	}
	COUNTRY_PATTERNS = {
					re.compile("[\w\s]+ cases [\w\s]+"): lambda country: data.get_country_data(country)['total_cases'],
                    re.compile("[\w\s]+ deaths [\w\s]+"): lambda country: data.get_country_data(country)['total_deaths'],                    
                    re.compile("[\w\s]+ recovered [\w\s]+"): lambda country: data.get_country_data(country)['total_recovered'],}
	result = None

	for pattern, func in COUNTRY_PATTERNS.items():
		if pattern.match(text):
			words = set(text.split(" "))
			for country in country_list:
				if country in words:
					result = func(country)
					break

	for pattern, func in TOTAL_PATTERNS.items():
		if pattern.match(text):
			result = func()
			break
	if result:
		assistant_speaks(result)

def audio_translator(src_lan, dest_lan):
    r = sr.Recognizer()
    translator = Translator()
    audio = ''
    with sr.Microphone() as source:
        audio = r.listen(source, phrase_time_limit=5)
        try:
        	text = r.recognize_google(audio, language=src_lan)
        	print("You: ", text.lower())
        	text_to_translate = translator.translate(text, src = src_lan, dest = dest_lan)
        	text = text_to_translate.text
        	print("Reply: ", text.lower())
        	speak = gTTS(text=text, lang = dest_lan, slow = False)
        	speak.save("captured_voice.mp3")
        	playsound.playsound("captured_voice.mp3")
        	os.remove("captured_voice.mp3")
        	return text.lower()
        except:
        	assistant_speaks("No Input Given!")
        	return

def lang_checker(check):
    r = sr.Recognizer()
    audio = ''
    with sr.Microphone() as source:
        audio = r.listen(source, phrase_time_limit=5)
        text = r.recognize_google(audio)
        print("You: ", text.lower())
        if text.lower() in check.values():
        	assistant_speaks(text.lower())
        	for key, value in check.items():
        		if text.lower() == value:
        			return key
        else:
        	assistant_speaks("Invalid Language!")
        	return False     	

def fun_translate():
	assistant_speaks("Provide Input Language: ")
	check = googletrans.LANGUAGES
	inp_lan = lang_checker(check)
	if inp_lan != False:
		assistant_speaks("Provide Output Language: ")
		oup_lan = lang_checker(check)
		if oup_lan != False:
			assistant_speaks("Speak...")
			audio_translator(inp_lan, oup_lan)
		else:
			return

	else:
		return
	return


if __name__ == "__main__":
    assistant_speaks("What's your name, Human?")
    name = get_audio()
    assistant_speaks("Hello, " + name + '.')
    while(1):
        assistant_speaks("What can i do for you?")
        text = get_audio()
        if text == 0:
        	continue
        assistant_speaks(text)
        if "exit" in str(text) or "bye" in str(text) or "go " in str(text) or "sleep" in str(text):
            assistant_speaks("Ok bye, "+ name+'.')
            break
        process_text(text)
        
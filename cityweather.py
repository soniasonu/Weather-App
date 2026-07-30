import requests
import json
import win32com.client as wincom
import os
from dotenv import load_dotenv

load_dotenv()

api_key = os.getenv("WEATHER_API_KEY")

city = input("Enter the name of the city: ")

url = f"https://api.weatherapi.com/v1/current.json?key={api_key}&q={city}, India"

r = requests.get(url)
wdic = json.loads(r.text)

temp_c = wdic["current"]["temp_c"]
humidity = wdic["current"]["humidity"]
condition = wdic["current"]["condition"]["text"]

print(f"Temperature: {temp_c}")
print(f"Humidity: {humidity}")
print(f"Condition: {condition}")

speak = wincom.Dispatch("SAPI.SpVoice")

text = (
    f"The current weather in {city} is {temp_c} degree Celsius. "
    f"Humidity is {humidity} percent. "
    f"The condition of the city is {condition}."
)

speak.Speak(text)
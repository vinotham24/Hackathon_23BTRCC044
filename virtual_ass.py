import os
import platform
import webbrowser
import datetime
import psutil
import requests
import pyttsx3
import speech_recognition as sr
import wmi
import json
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
from comtypes import CLSCTX_ALL

# Initialize Text-to-Speech Engine
engine = pyttsx3.init()

# File to store user preferences
PREFERENCES_FILE = "user_preferences.json"

def save_preferences(data):
    """Save user preferences to a file."""
    with open(PREFERENCES_FILE, "w") as file:
        json.dump(data, file)

def load_preferences():
    """Load user preferences from a file."""
    if os.path.exists(PREFERENCES_FILE):
        with open(PREFERENCES_FILE, "r") as file:
            return json.load(file)
    return {}

# Speak function
def speak(text):
    """Convert text to speech."""
    engine.say(text)
    engine.runAndWait()

# Function to listen to voice commands
def listen():
    """Listen for user voice input."""
    recognizer = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        recognizer.adjust_for_ambient_noise(source)
        try:
            audio = recognizer.listen(source, timeout=5)
            command = recognizer.recognize_google(audio).lower()
            print(f"You said: {command}")
            return command
        except sr.UnknownValueError:
            speak("Sorry, I didn't understand.")
            return None
        except sr.RequestError:
            speak("Could not connect to the internet.")
            return None
        except sr.WaitTimeoutError:
            speak("You didn't say anything.")
            return None

# Get current date
def get_date():
    return datetime.date.today().strftime("%A, %B %d, %Y")

# Get current time
def get_time():
    return datetime.datetime.now().strftime("%I:%M %p")

# Get battery status
def get_battery_status():
    battery = psutil.sensors_battery()
    return f"Battery is at {battery.percent}%" if battery else "Battery status not available"

# Dynamic greetings based on time of day
def get_greeting():
    hour = datetime.datetime.now().hour
    if hour < 12:
        return "Good morning!"
    elif 12 <= hour < 18:
        return "Good afternoon!"
    else:
        return "Good evening!"

# Open system apps
def open_app(app_name):
    system = platform.system().lower()
    apps = {
        "windows": {
            "spotify": "start spotify",
            "chrome": "start chrome",
            "calculator": "calc",
            "notepad": "notepad",
            "settings": "start ms-settings:",
            "file explorer": "explorer"
        },
        "linux": {
            "spotify": "spotify",
            "chrome": "google-chrome",
            "calculator": "gnome-calculator",
            "notepad": "gedit",
            "settings": "gnome-control-center",
            "file explorer": "nautilus"
        },
        "darwin": {  # macOS
            "spotify": "open -a Spotify",
            "chrome": "open -a Google\\ Chrome",
            "calculator": "open -a Calculator",
            "notepad": "open -a TextEdit",
            "settings": "open -a System\\ Preferences",
            "file explorer": "open -a Finder"
        }
    }

    if system in apps and app_name in apps[system]:
        os.system(apps[system][app_name])
        speak(f"Opening {app_name}")
    else:
        websites = {
            "youtube": "https://www.youtube.com",
            "whatsapp": "https://web.whatsapp.com",
            "instagram": "https://www.instagram.com",
            "twitter": "https://twitter.com",
            "linkedin": "https://www.linkedin.com"
        }
        if app_name in websites:
            webbrowser.open(websites[app_name])
            speak(f"Opening {app_name}")
        else:
            speak("App or website not found.")

# Play a song on Spotify first, then YouTube if Spotify is not installed
def play_song(song_name):
    query = song_name.replace("play ", "").strip()
    system = platform.system().lower()
    spotify_cmds = {"windows": "start spotify", "linux": "spotify", "darwin": "open -a Spotify"}
    
    if system in spotify_cmds:
        try:
            os.system(spotify_cmds[system])
            speak(f"Playing {query} on Spotify.")
        except Exception:
            speak("Spotify not available. Playing on YouTube instead.")
            webbrowser.open(f"https://www.youtube.com/results?search_query={query}")
    else:
        webbrowser.open(f"https://www.youtube.com/results?search_query={query}")
        speak(f"Playing {query} on YouTube.")

# Lock the screen
def lock_screen():
    os.system("rundll32.exe user32.dll,LockWorkStation")
    speak("Locking your screen")

# Get weather using wttr.in
def get_weather(city):
    url = f"https://wttr.in/{city}?format=%C+%t"
    try:
        response = requests.get(url)
        return f"Weather in {city}: {response.text.strip()}" if response.status_code == 200 else f"Could not fetch weather for {city}."
    except requests.exceptions.RequestException:
        return "Failed to fetch weather data."

# Adjust brightness
def set_brightness(level):
    if 0 <= level <= 100:
        w = wmi.WMI(namespace='wmi')
        methods = w.WmiMonitorBrightnessMethods()[0]
        methods.WmiSetBrightness(level, 0)
        speak(f"Brightness set to {level}%")
    else:
        speak("Brightness must be between 0 and 100.")

# Adjust volume
def change_volume(level):
    if 0 <= level <= 100:
        devices = AudioUtilities.GetSpeakers()
        interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
        volume = interface.QueryInterface(IAudioEndpointVolume)
        volume.SetMasterVolumeLevelScalar(level / 100, None)
        speak(f"Volume set to {level}%")

# System control (Shutdown, Restart, Sleep)
def system_control(action):
    commands = {"shutdown": "shutdown /s /t 5", "restart": "shutdown /r /t 5", "sleep": "rundll32.exe powrprof.dll,SetSuspendState 0,1,0"}
    if action in commands:
        os.system(commands[action])
        speak(f"System {action} initiated")

# Main loop
def main():
    preferences = load_preferences()
    
    if "username" not in preferences:
        speak("Can you please enter your name?")
        preferences["username"] = input("What is your name? ").strip()
        save_preferences(preferences)
        speak(f"Nice to meet you, {preferences['username']}! {get_greeting()} How can I assist you today?")
    else:
        speak(f"{get_greeting()} Welcome back, {preferences['username']}!")
    
    while True:
        command = listen()
        if command:
            if "date" in command:
                speak(f"Today's date is {get_date()}")

            elif "time" in command:
                speak(f"The current time is {get_time()}")

            elif "battery" in command:
                speak(get_battery_status())

            elif "open" in command:
                open_app(command.replace("open ", ""))

            elif "play" in command:
                play_song(command)

            elif "lock screen" in command:
                lock_screen()

            elif "weather in" in command:
                city = command.replace("weather in", "").strip()
                speak(get_weather(city))
                
            elif "brightness" in command:
                try:
                    level = int(''.join(filter(str.isdigit, command)))
                    set_brightness(level)
                except ValueError:
                    speak("Please specify a brightness level between 0 and 100.")

            elif "volume" in command:
                level = int(''.join(filter(str.isdigit, command)))
                change_volume(level)

            elif "shutdown" in command:
                system_control("shutdown")

            elif "restart" in command:
                system_control("restart")

            elif "sleep" in command:
                system_control("sleep")

            elif "exit" in command or "goodbye" in command or "bye" in command:
                speak(f"Goodbye! Have a great day, {preferences['username']}!")
                break
            else:
                speak("Sorry, I didn't understand that command.")

if __name__ == "__main__":
    main()
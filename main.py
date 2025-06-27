import pyttsx3
import speech_recognition as sr
import random
import webbrowser
import datetime
import os
import time
from plyer import notification
import pyautogui
import sys
from PIL import ImageGrab
import keyboard
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
import wikipedia
import pywhatkit as pwk
import smtplib,ssl
import requests
import mtranslate

# Initialize TTS engine with Indian English voice
engine = pyttsx3.init()
voices = engine.getProperty('voices')
indian_voice = None
for voice in voices:
    if "indian" in voice.name.lower():
        indian_voice = voice.id
        break
engine.setProperty('voice', indian_voice if indian_voice else voices[0].id)
engine.setProperty('rate', 145)
engine.setProperty('volume', 1.0)

# Initialize audio controller
devices = AudioUtilities.GetSpeakers()
interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
volume = cast(interface, POINTER(IAudioEndpointVolume))

# Initialize recognizer with Indian English
r = sr.Recognizer()
r.energy_threshold = 4000
r.dynamic_energy_threshold = True

def speak(command):
    engine.say(command)
    engine.runAndWait()

def listen_for_text():
    """
    Captures audio input from the user and converts it to text using Google's Speech Recognition API.
    """
    with sr.Microphone() as source:
        print("Listening...")
        r.adjust_for_ambient_noise(source, duration=0.5)
        try:
            audio = r.listen(source, timeout=10, phrase_time_limit=20)  # Extended time for conversational input
            text = r.recognize_google(audio, language='en-IN')
            return text.lower()
        except sr.UnknownValueError:
            speak("I couldn't understand. Please try again.")
            return None
        except sr.RequestError:
            speak("Network error occurred.")
            return None


def control_system_volume(command):
    try:
        current_volume = round(volume.GetMasterVolumeLevelScalar() * 100)
        
        if "set volume" in command:
            level = int(''.join(filter(str.isdigit, command)))
            if 0 <= level <= 100:
                volume.SetMasterVolumeLevelScalar(level / 100, None)
                speak(f"Volume set to {level} percent")
            
        elif "increase volume" in command:
            new_volume = min(current_volume + 10, 100)
            volume.SetMasterVolumeLevelScalar(new_volume / 100, None)
            speak(f"Volume increased to {new_volume} percent")
            
        elif "decrease volume" in command:
            new_volume = max(current_volume - 10, 0)
            volume.SetMasterVolumeLevelScalar(new_volume / 100, None)
            speak(f"Volume decreased to {new_volume} percent")
            
        elif "mute" in command:
            volume.SetMute(1, None)
            speak("Audio muted")
            
        elif "unmute" in command:
            volume.SetMute(0, None)
            speak("Audio unmuted")
    except Exception as e:
        speak(f"Error controlling volume: {str(e)}")

def control_window(command):
    try:
        if "minimise" in command:
            pyautogui.hotkey('win', 'down')
            speak("Window minimized")
        elif "maximize" in command:
            pyautogui.hotkey('win', 'up')
            speak("Window maximized")
        elif "restore" in command:
            pyautogui.hotkey('win', 'down')
            speak("Window restored")
    except Exception as e:
        speak(f"Error controlling window: {str(e)}")

def get_chat_response(user_input):
    """
    Get a response from the Hugging Face conversational model.
    """
    import requests

    # Hugging Face API Endpoint and Token
    HUGGING_FACE_API_URL = "https://api-inference.huggingface.co/models/facebook/blenderbot-400M-distill"
    HUGGING_FACE_API_TOKEN = "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"  # Replace with your Hugging Face token

    headers = {"Authorization": f"Bearer {HUGGING_FACE_API_TOKEN}"}

    try:
        # Payload for the model
        payload = {"inputs": {"text": user_input}}

        # Make the API request
        response = requests.post(HUGGING_FACE_API_URL, headers=headers, json=payload)
        response.raise_for_status()  # Raise an error for HTTP issues
        data = response.json()

        # Extract and return the generated response
        if isinstance(data, dict) and "generated_text" in data:
            return data["generated_text"]
        elif isinstance(data, list) and len(data) > 0:
            return data[0].get("generated_text", "Sorry, I couldn't generate a response.")
        else:
            return "Sorry, I couldn't generate a response."
    except requests.exceptions.RequestException as e:
        print(f"Error fetching response from Hugging Face: {e}")
        return "An error occurred while fetching the response."


def whatsapp_send_message_background():
    """Send a WhatsApp message using the installed app without displaying the window."""
    try:
        # Open WhatsApp Desktop minimized
        speak("Opening WhatsApp.")
        os.system("start /min whatsapp://")  # Opens WhatsApp in minimized mode
        time.sleep(5)  # Wait for WhatsApp to initialize

        # Ask for the contact name
        speak("Whom should I send the message to?")
        contact_name = listen_for_text()
        if not contact_name:
            speak("I couldn't hear the contact name. Please try again.")
            return

        # Search for the contact
        pyautogui.hotkey("ctrl", "f")  # Open the search bar in WhatsApp
        time.sleep(1)
        pyautogui.write(contact_name, interval=0.1)  # Type the contact name
        time.sleep(2)
        pyautogui.press("enter")  # Select the contact
        time.sleep(1)

        # Ask for the message
        speak(f"What message should I send to {contact_name}?")
        message = listen_for_text()
        if not message:
            speak("I couldn't hear the message. Please try again.")
            return

        # Type and send the message
        pyautogui.write(message, interval=0.1)  # Type the message
        pyautogui.press("enter")  # Send the message
        speak("Message sent successfully.")

    except Exception as e:
        speak(f"An error occurred while sending the message: {str(e)}")
        print(f"Error: {e}")



def media_control(command):
    try:
        if "pause" in command:
            keyboard.press_and_release('space')
            speak("Media paused")
        elif "play" in command:
            keyboard.press_and_release('space')
            speak("Media playing")
        elif "next" in command:
            keyboard.press_and_release('shift+n')
            speak("Next media")
        elif "previous" in command:
            keyboard.press_and_release('shift+p')
            speak("Previous media")
        elif "forward" in command:
            keyboard.press_and_release('right')
            speak("Forwarded")
        elif "rewind" in command:
            keyboard.press_and_release('left')
            speak("Rewinded")
    except Exception as e:
        speak(f"Error controlling media: {str(e)}")

def improved_typing(text):
    try:
        keyboard.press_and_release('ctrl+a')
        keyboard.press_and_release('backspace')
        time.sleep(0.5)
        
        words = text.split()
        for word in words:
            pyautogui.write(word)
            pyautogui.press('space')
            time.sleep(0.1)
            
        speak("Text has been written")
    except Exception as e:
        speak(f"Error typing text: {str(e)}")

def perform_search(app_context=None):
    speak("What would you like to search for?")
    search_query = listen_for_text()
    
    if search_query:
        if app_context == "youtube":
            url = f"https://www.youtube.com/results?search_query={search_query}"
            webbrowser.open(url)
            time.sleep(2)
            pyautogui.press('tab', presses=3)
            pyautogui.press('enter')
        elif app_context == "google":
            webbrowser.open(f"https://www.google.com/search?q={search_query}")
            speak(f"Searching for {search_query} on Google")
        elif app_context == "wikipedia":
            try:
                summary = wikipedia.summary(search_query, sentences=2)
                speak(f"According to Wikipedia: {summary}")
                print(summary)
            except wikipedia.exceptions.DisambiguationError:
                speak("There are multiple results for your query. Please be more specific.")
            except wikipedia.exceptions.PageError:
                speak("No results found on Wikipedia.")
            except Exception as e:
                speak(f"An error occurred: {str(e)}")
        elif app_context == "whatsapp":
            try:
                speak("Please provide the phone number to send the message.")
                phone_number = listen_for_text().replace(" ", "")
                speak("What message would you like to send?")
                message = listen_for_text()
                if phone_number and message:
                    pwk.sendwhatmsg_instantly(f"+{+917892787731}", message, wait_time=10)
                    speak("Message sent successfully.")
                else:
                    speak("Failed to capture the phone number or message.")
            except Exception as e:
                speak(f"An error occurred: {str(e)}")
        else:
            improved_typing(search_query)
            pyautogui.press('enter')
            speak(f"Searching for {search_query}")


def take_screenshot():
    try:
        if not os.path.exists("screenshots"):
            os.makedirs("screenshots")
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"screenshots/screenshot_{timestamp}.png"
        screenshot = ImageGrab.grab()
        screenshot.save(filename)
        speak(f"Screenshot has been saved")
    except Exception as e:
        speak(f"Error taking screenshot: {str(e)}")

# Add the new function here
from huggingface_hub import InferenceClient

# Initialize Hugging Face client with your API key
client = InferenceClient(
    model="stabilityai/stable-diffusion-3.5-large", 
    token="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"  # Hugging face api access key
)

def generate_image_from_text(text_description):
    """Generates an image using Hugging Face's Stable Diffusion model."""
    try:
        # Generate the image
        speak("Generating your image, please wait...")
        image = client.text_to_image(text_description)  # Output is a PIL.Image object
        
        # Save the image locally
        if not os.path.exists("generated_images"):
            os.makedirs("generated_images")
        image_path = f"generated_images/{text_description.replace(' ', '_')}.png"
        image.save(image_path)
        
        # Provide feedback to the user
        speak("Image generated successfully.")
        print(f"Image saved at: {image_path}")
    except Exception as e:
        speak("An error occurred while generating the image.")
        print(f"Error: {e}")

def handle_command(command):
    """Handles voice commands."""
    if "generate image" in command:
        speak("What should the image be about?")
        try:
            # Capture user input for the image description
            text_description = listen_for_text()

            # Check if a valid description was provided
            if text_description:
                generate_image_from_text(text_description)
            else:
                speak("I didn't catch that. Please describe the image again.")
        except Exception as e:
            # Handle unexpected errors gracefully
            speak(f"An unexpected error occurred: {str(e)}")
        return True
    
    if "chat" in command or "talk" in command:
        speak("What would you like to discuss?")
        while True:
            user_input = listen_for_text()
            if not user_input or user_input in ["exit", "quit", "stop"]:
                speak("Exiting chat mode. Let me know if there's anything else.")
                break
            response = get_chat_response(user_input)
            print(f"ChatGPT-like model response: {response}")
            speak(response)
        return True

    # Other command handlers go here...

def open_application(query):
    try:
        pyautogui.press("win")
        time.sleep(1)
        pyautogui.write(query, interval=0.1)
        time.sleep(1)
        pyautogui.press("enter")
        speak(f"Opening {query}")
    except Exception as e:
        speak(f"Error opening application: {str(e)}")

def close_specific_app(app_name):
    try:
        app_processes = {
            'chrome': 'chrome.exe',
            'firefox': 'firefox.exe',
            'excel': 'excel.exe',
            'word': 'winword.exe',
            'notepad': 'notepad.exe',
            'spotify': 'spotify.exe',
            'discord': 'discord.exe',
            'steam': 'steam.exe',
            'telegram': 'telegram.exe',
            'whatsapp': 'whatsapp.exe',
            'opera gx': 'opera.exe',
            'youtube': 'chrome.exe',
            'instagram': 'chrome.exe',
            'lenovo vantage': 'LenovoVantage.exe',
            'microsoft store': 'WinStore.App.exe'
        }
        
        process_name = app_processes.get(app_name.lower())
        if process_name:
            os.system(f'taskkill /IM {process_name} /F')
            speak(f"Closing {app_name}")
        else:
            speak(f"Sorry, I don't know how to close {app_name}")
    except Exception as e:
        speak(f"Unable to close {app_name}: {str(e)}")

def system_power_control(action):
    try:
        if action == "shutdown":
            speak("Shutting down the system")
            os.system("shutdown /s /t 5")
        elif action == "restart":
            speak("Restarting the system")
            os.system("shutdown /r /t 5")
        elif action == "cancel":
            speak("Canceling shutdown or restart")
            os.system("shutdown /a")
    except Exception as e:
        speak(f"Error controlling system power: {str(e)}")

def add_task(command):
    task = command.replace("new task", "").strip()
    if task:
        speak(f"Adding task: {task}")
        with open("todo.txt", "a") as file:
            file.write(task + "\n")
        print(f"Task added: {task}")
    else:
        speak("You didn't specify a task. Please try again.")

def read_tasks():
    try:
        with open("todo.txt", "r") as file:
            tasks = file.readlines()
        
        if tasks:
            speak("Here are your tasks.")
            task_list = [f"Task {i + 1}: {task.strip()}" for i, task in enumerate(tasks)]
            for task in task_list:
                speak(task)
                print(task)
        else:
            speak("Your to-do list is empty.")
    except FileNotFoundError:
        speak("You don't have a to-do list yet.")

def show_work():
    try:
        with open("todo.txt", "r") as file:
            tasks = file.read()
        notification.notify(
            title="Today's work",
            message=tasks
        )
        speak("Displaying today's work")
    except FileNotFoundError:
        speak("No todo list found")

def get_current_time():
    now = datetime.datetime.now()
    return f"The current time is {now.strftime('%I:%M %p')}."

def get_current_date():
    today = datetime.datetime.now()
    return f"Today's date is {today.strftime('%A, %d %B %Y')}."

def handle_command(command):
    """Handles voice commands."""
    commands = {
        "hello": "Welcome, how can I help you?",
        "how are you": "I'm just a program, but I'm here to help you!",
        "what is your name": "My name is NOVA AI.",
        "who is your creator": "My creator is Siddu Ambesange",
        "say the time": get_current_time(),
        "say the date": get_current_date(),
        "exit": "Goodbye! Have a great day.",
        "quit": "Goodbye! Have a great day.",
        "stop": "Goodbye! Have a great day."
    }

    # Check if the command matches any predefined ones
    for key, response in commands.items():
        if key in command:
            print(f"NOVA AI responds: {response}")
            speak(response)
            if key in ["exit", "quit", "stop"]:
                sys.exit()
            return True  # Continue listening for more commands

    # Handle image generation command if it's not in the predefined list
    if "generate image" in command:
        speak("What should the image be about?")
        try:
            # Attempt to capture user input
            text_description = listen_for_text()

            # Check if a valid description was provided
            if text_description:
                generate_image_from_text(text_description)
            else:
                speak("I didn't catch that. Please describe the image again.")
        except Exception as e:
            # Handle unexpected errors gracefully
            speak(f"An error occurred: {e}")
        return True

    # Music handling
    if "play music" in command:
        speak("Playing music")
        song = random.randint(1, 3)
        if song == 1:
            webbrowser.open("https://www.youtube.com/watch?v=p7743kxiMao&list=RDp7743kxiMao&start_radio=1")
        elif song == 2:
            webbrowser.open("https://www.youtube.com/watch?v=BOBrgRoR-pM&list=RDp7743kxiMao&index=8")
        elif song == 3:
            webbrowser.open("https://www.youtube.com/watch?v=TeThp6h1K3w&list=RDp7743kxiMao&index=14")
        return True

    # Further command handling can continue below...


    if "open" in command:
        app_name = command.replace("open", "").strip()
        if app_name:
            open_application(app_name)
            return True

    if "close" in command:
        for app in ['whatsapp', 'chrome', 'notepad', 'word', 'excel', 'telegram', 'spotify', 
                'microsoft store', 'discord', 'lenovo vantage', 'steam', 'instagram', 
                'opera gx', 'youtube', 'firefox','downloads','file explorer']:
            if app in command:
                close_specific_app(app)
                return True

    if "search" in command:
        if "youtube" in command:
            perform_search("youtube")
        elif "google" in command:
            perform_search("google")
        elif "wikipedia" in command:
            perform_search("wikipedia")
        else:
            perform_search()
        return True

    if "type" in command:
        speak("What would you like me to write?")
        text = listen_for_text()
        if text:
            improved_typing(text)
        return True
    
    if "whatsapp" in command:
        whatsapp_send_message_background()
        return True

    if "screenshot" in command:
        take_screenshot()
        return True

    if "show work" in command:
        show_work()
        return True

    if "new task" in command:
        add_task(command)
        return True

    if "read task" in command or "read my to-do list" in command:
        read_tasks()
        return True

    if "shutdown system" in command or "power off" in command:
        system_power_control("shutdown")
        return True
        
    if "restart system" in command or "reboot" in command:
        system_power_control("restart")
        return True
        
    if "cancel shutdown" in command or "cancel restart" in command:
        system_power_control("cancel")
        return True

    if any(phrase in command for phrase in ["volume", "mute", "unmute"]):
        control_system_volume(command)
        return True
        
    if any(phrase in command for phrase in ["minimize", "maximize", "restore"]):
        control_window(command)
        return True
        
    if any(phrase in command for phrase in ["pause", "play", "next", "previous", "forward", "rewind"]):
        media_control(command)
        return True

    return False


def main():
    # Update these parameters at the start of your main function
    r.energy_threshold = 3000
    r.dynamic_energy_threshold = True
    r.dynamic_energy_adjustment_damping = 0.15
    r.dynamic_energy_ratio = 1.5
    r.pause_threshold = 0.6
    r.phrase_threshold = 0.3
    
    print("Adjusting for ambient noise, please wait...")
    with sr.Microphone() as source:
        r.adjust_for_ambient_noise(source, duration=0.5)
        print("Listening is ready. Speak when you are ready.")
        speak("Starting up the nova ai.")
        
        while True:
            try:
                print("Waiting for your voice...")
                audio = r.listen(source, timeout=10)
                print("Audio captured successfully.")
                
                recognized_text = r.recognize_google(audio, language='en-IN')
                recognized_text = recognized_text.lower()
                print(f"NOVA AI thinks you said: {recognized_text}")
                
                if not handle_command(recognized_text):
                    speak("Please try again.")
            except sr.UnknownValueError:
                print("Could not understand. Try again.")
            except sr.WaitTimeoutError:
                print("No speech detected. Still waiting...")
            except sr.RequestError as e:
                print(f"Could not request results from Google Speech Recognition service; {e}")
                speak("There was an issue with the speech recognition service.")
            except Exception as e:
                print(f"An unexpected error occurred: {e}")
                speak("An unexpected error occurred while processing your speech.")

if __name__ == "__main__":
    main()
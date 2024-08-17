import os
import sys
import tkinter as tk
from tkinter import scrolledtext
from tkinter import ttk
import threading
import subprocess
import webbrowser

# Add the path where google-generativeai is installed
sys.path.append(r'C:\Users\Salman\AppData\Roaming\Python\Python312\site-packages')
import speech_recognition as sr
import pyttsx3
import google.generativeai as genai

# Path for Chrome (update this path based on your system's Chrome installation)
chrome_path = "C:/Program Files/Google/Chrome/Application/chrome.exe %s"

# Initialize the recognizer and text-to-speech engine
recognizer = sr.Recognizer()
tts_engine = pyttsx3.init()

# Ensure the API key is set for GUB AI Chat
api_key = os.getenv("GENAI_API_KEY", "AIzaSyBvmi_-eRZDHfxEtaJd40RrJTDmBVZhJ_c")
genai.configure(api_key=api_key)

# Create the model for GUB AI Chat
generation_config = {
    "temperature": 1,
    "top_p": 0.95,
    "top_k": 64,
    "max_output_tokens": 8192,
    "response_mime_type": "text/plain",
}

model = genai.GenerativeModel(
    model_name="gemini-1.5-flash",
    generation_config=generation_config,
)

chat_session = model.start_chat(history=[])

# Flags to control the assistant's listening state
listening = False
running = True

# Function to speak text
def speak(text):
    tts_engine.say(text)
    tts_engine.runAndWait()

# Function to listen to commands
def listen():
    with sr.Microphone() as source:
        recognizer.adjust_for_ambient_noise(source)
        update_status("Listening...")
        audio = recognizer.listen(source)
        try:
            command = recognizer.recognize_google(audio)
            update_status(f"You said: {command}")
            return command.lower()
        except sr.UnknownValueError:
            update_status("Sorry, I did not understand that.")
            speak("Sorry, I did not understand that.")
            return ""
        except sr.RequestError:
            update_status("Sorry, my speech service is down.")
            speak("Sorry, my speech service is down.")
            return ""

# Function to open applications
def open_application(app_name):
    app_paths = {
        "chrome": "C:/Program Files/Google/Chrome/Application/chrome.exe",
        "ms word": "C:/Program Files/Microsoft Office/root/Office16/WINWORD.EXE",
        "excel": "C:/Program Files/Microsoft Office/root/Office16/EXCEL.EXE",
        "onenote": "C:/Program Files/Microsoft Office/root/Office16/ONENOTE.EXE",
        "edge": "C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe",
        "powerpoint": "C:/Program Files/Microsoft Office/root/Office16/POWERPNT.EXE",
        "winrar": "C:/Program Files/WinRAR/WinRAR.exe",
        "vs code": "C:/Users/Salman/AppData/Local/Programs/Microsoft VS Code/Code.exe"
    }
    
    if app_name in app_paths:
        try:
            subprocess.Popen([app_paths[app_name]])
            speak(f"Opening {app_name}")
            update_status(f"Opening {app_name}")
        except FileNotFoundError:
            speak(f"Sorry, I couldn't find the application {app_name}")
            update_status(f"Sorry, I couldn't find the application {app_name}")
    else:
        speak(f"Application {app_name} is not configured.")
        update_status(f"Application {app_name} is not configured.")

# Function to open websites with Google Chrome
def open_website(url, action_message="Opening URL"):
    try:
        update_status(f"{action_message}: {url}")
        webbrowser.get(chrome_path).open(url)
    except webbrowser.Error:
        speak("Failed to open browser.")
        update_status("Failed to open browser.")

# Function to process voice commands
def process_command(command):
    if "open" in command:
        command = command.replace("open", "").strip()
        if command in ["chrome", "ms word", "excel", "onenote", "edge", "powerpoint", "winrar", "vs code"]:
            open_application(command)
        elif command in ["youtube", "facebook", "instagram"]:
            open_website(f"https://www.{command}.com", f"Opening {command}")
        else:
            speak(f"Command not recognized: {command}. Please specify a valid application or website.")
            update_status(f"Command not recognized: {command}. Please specify a valid application or website.")
    elif "search" in command:
        search_query = command.replace("search", "").strip()
        if search_query:
            google_search_url = f"https://www.google.com/search?q={search_query.replace(' ', '+')}"
            open_website(google_search_url, f"Searching {search_query} for you")
            speak(f"Searching {search_query} for you")
        else:
            speak("Please specify what you want to search for.")
            update_status("Please specify what you want to search for.")
    else:
        speak("I can only open websites, applications, or search in Google for now.")
        update_status("I can only open websites, applications, or search in Google for now.")

# Function to start listening for commands
def start_listening():
    global listening, running
    speak("Hello, how can I assist you today?")
    update_status("Hello, how can I assist you today?")
    while running:
        if listening:
            command = listen()
            if command:
                if "exit" in command or "quit" in command or "stop" in command:
                    speak("Goodbye!")
                    update_status("Goodbye!")
                    break
                process_command(command)

# Function to update status in the text area
def update_status(message):
    status_text.config(state=tk.NORMAL)
    status_text.insert(tk.END, message + '\n')
    status_text.config(state=tk.DISABLED)
    status_text.see(tk.END)

# Function to start the assistant in a separate thread
def start_assistant():
    global listening, running
    listening = True
    running = True
    threading.Thread(target=start_listening).start()

# Function to pause the assistant
def pause_assistant():
    global listening
    listening = False
    speak("Assistant paused.")
    update_status("Assistant paused.")

# Function to stop the assistant
def stop_assistant():
    global running
    running = False
    speak("Assistant stopped.")
    update_status("Assistant stopped.")
    root.quit()  # Close the GUI

# Function to resume the assistant
def resume_assistant():
    global listening
    listening = True
    speak("Assistant resumed.")
    update_status("Assistant resumed.")

# Function to send a message to GUB AI Chat
def send_message(event=None):
    user_input = user_input_field.get().strip()  # Get user input and strip any extra whitespace
    if user_input.lower() in ['exit', 'quit']:
        output_text.insert(tk.END, "Ending chat session.\n")
        return
    
    response = chat_session.send_message(user_input)
    output_text.insert(tk.END, f"You: {user_input}\n")
    output_text.insert(tk.END, f"GUB Assistant: {response.text}\n")
    user_input_field.delete(0, tk.END)

# GUI setup
root = tk.Tk()
root.title("GUB AI Assistant")
root.geometry("550x700")
root.configure(bg="#28a745")  # Set background color to greenish (#72cf94)

# Style configuration with greenish theme
style = ttk.Style()
style.configure("TFrame", background="#28a745")
style.configure("TLabel", background="#28a745", foreground="white", font=("Helvetica", 20, "bold"))  # Set foreground to white and make it bold
style.configure("TButton", font=("Helvetica", 10), foreground="black", background="#4CAF50")  # Green button
style.map("TButton", background=[('active', '#45a049')])  # Change color on hover
style.configure("TEntry", font=("Helvetica", 10))
style.configure("TScrolledText", font=("Helvetica", 10))

# Frame for Assistant controls
control_frame = ttk.Frame(root)
control_frame.pack(pady=10)

start_button = ttk.Button(control_frame, text="Start Listening", command=start_assistant)
start_button.grid(row=0, column=0, padx=5, pady=5)

pause_button = ttk.Button(control_frame, text="Pause", command=pause_assistant)
pause_button.grid(row=0, column=1, padx=5, pady=5)

stop_button = ttk.Button(control_frame, text="Stop", command=stop_assistant)
stop_button.grid(row=0, column=2, padx=5, pady=5)

resume_button = ttk.Button(control_frame, text="Resume", command=resume_assistant)
resume_button.grid(row=0, column=3, padx=5, pady=5)

# Frame for status updates (Voice Assistant's output)
status_frame = ttk.Frame(root)
status_frame.pack(pady=10)

status_text = scrolledtext.ScrolledText(status_frame, wrap=tk.WORD, height=10, width=60, state=tk.DISABLED, font=("Helvetica", 10))
status_text.pack()

# Frame for GUB AI Chat
chat_frame = ttk.Frame(root)
chat_frame.pack(pady=10)

# Label for "Chat With GUB AI" centered on one line with reduced padding
user_input_label = ttk.Label(chat_frame, text="Chat With GUB AI:", foreground="white", font=("Helvetica", 20, "bold"))
user_input_label.grid(row=0, column=0, pady=1, sticky='n')  # Reduce pady to 5 for reduced spacing

# Frame for user input and submit button on the next line
input_frame = ttk.Frame(chat_frame)
input_frame.grid(row=1, column=0, padx=10, pady=10)

user_input_field = ttk.Entry(input_frame, width=40)
user_input_field.grid(row=0, column=0, padx=5, pady=5)

send_button = ttk.Button(input_frame, text="Send", command=send_message)
send_button.grid(row=0, column=1, padx=5, pady=5)

# Bind Enter key to send_message function
user_input_field.bind('<Return>', send_message)

# Output text box for Chat with GUB AI
output_text = scrolledtext.ScrolledText(chat_frame, wrap=tk.WORD, height=15, width=60, state=tk.NORMAL, font=("Helvetica", 10))
output_text.grid(row=2, column=0, padx=10, pady=10)

# Start the GUI main loop
root.mainloop()

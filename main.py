# Importing the modules
from ctypes.wintypes import SERVICE_STATUS_HANDLE
import smtplib
import speech_recognition as sr
import pyttsx3
from email.message import EmailMessage
from tkinter import *
from tkinter import messagebox
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import os
import pickle
import threading

# Creating the GUI window
window = Tk()
window.title("Voice Based Email Sender")
window.geometry("350x700")
# window.configure(bg="#f0f0f0")

# Creating a function to round the corners of the window
def round_corners(window):
    window.overrideredirect(True)
    window.geometry("+550+200")
    window.wm_attributes("-transparentcolor", "white")
    canvas = Canvas(window, width=350, height=700, bg="white", highlightthickness=0)
    canvas.pack()
    canvas.create_arc(0, 0, 20, 20, start=90, extent=90, fill="black")
    canvas.create_arc(330, 0, 350, 20, start=0, extent=90, fill="black")
    canvas.create_arc(0, 680, 20, 700, start=180, extent=90, fill="black")
    canvas.create_arc(330, 680, 350, 700, start=270, extent=90, fill="black")
    canvas.create_rectangle(10, 0, 340, 700, fill="black")
    canvas.create_rectangle(0, 10, 350, 690, fill="black")
    return canvas

# Calling the round corners function
canvas = round_corners(window)

# Creating the sender label
sender_label = Label(window, text="Sender:", bg="#f0f0f0", font=("Arial", 12))
canvas.create_window(175, 35, window=sender_label)

# Creating the inbox button
inbox_button = Button(window, text="Inbox", command=lambda: show_inbox(), bg="#00ffff", font=("Arial", 12))
canvas.create_window(175, 85, window=inbox_button)

# Creating the compose button
compose_button = Button(window, text="Compose", command=lambda: show_compose(), bg="#2C2C2C", font=("Arial", 12))
canvas.create_window(175, 135, window=compose_button)

# Creating the speech recognition function
def recognize_speech():
    # Initializing the speech recognizer and engine
    recognizer = sr.Recognizer()
    engine = pyttsx3.init()
    voices = engine.getProperty("voices")
    engine.setProperty("voice", voices[1].id)
    rate = engine.getProperty("rate")
    engine.setProperty("rate", rate-20)
    
    # Listening to the microphone input
    with sr.Microphone() as source:
        recognizer.pause_threshold = 1
        recognizer.adjust_for_ambient_noise(source, duration=1)
        engine.say("Please speak now")
        engine.runAndWait()
        audio = recognizer.listen(source)
    
    # Converting the audio to text
    try:
        text = recognizer.recognize_google(audio)
        print(text)
        engine.say("You said: " + text)
        engine.runAndWait()
        return text
    except:
        engine.say("Sorry, I could not understand that")
        engine.runAndWait()
        return ""

# Creating the authenticate gmail function
def authenticate_gmail():
    # Defining the scopes
    SCOPES = ["https://www.googleapis.com/auth/gmail.readonly", "https://www.googleapis.com/auth/gmail.send"]
    
    # Checking the credentials
    creds = None
    if os.path.exists("token.pickle"):
        with open("token.pickle", "rb") as token:
            creds = pickle.load(token)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
            creds = flow.run_local_server(port=0)
        with open("token.pickle", "wb") as token:
            pickle.dump(creds, token)
    
    # Building the service
    service = build("gmail", "v1", credentials=creds)
    return service

# Creating the show inbox function
def show_inbox():
    # Creating a new window
    inbox_window = Toplevel(window)
    inbox_window.title("Inbox")
    inbox_window.geometry("350x700")
    inbox_window.configure(bg="#f0f0f0")
    
    # Creating a loading label
    loading_label = Label(inbox_window, text="Loading...", bg="#f0f0f0", font=("Arial", 12))
    loading_label.pack(pady=10)
    
    # Creating a thread to load the emails
    def load_emails():
        # Authenticating gmail
        service = authenticate_gmail()
        
        # Getting the user's email address
        user_profile = service.users().getProfile(userId="me").execute()
        user_email = user_profile["emailAddress"]
        
        # Displaying the user's email address
        user_label = Label(inbox_window, text="User: " + user_email, bg="#f0f0f0", font=("Arial", 12))
        user_label.pack(pady=10)
        
        # Getting the unread messages
        results = service.users().messages().list(userId="me", labelIds=["INBOX", "UNREAD"]).execute()
        messages = results.get("messages", [])
        
        # Displaying the unread messages in a listview with scrollbar
        list_frame = Frame(inbox_window)
        list_frame.pack(fill=BOTH, expand=1)
        list_scroll = Scrollbar(list_frame)
        list_scroll.pack(side=RIGHT, fill=Y)
        list_view = Listbox(list_frame, yscrollcommand=list_scroll.set)
        list_view.pack(fill=BOTH, expand=1)
        list_scroll.config(command=list_view.yview)
        if messages:
            for message in messages:
                msg = service.users().messages().get(userId="me", id=message["id"]).execute()
                headers = msg["payload"]["headers"]
                subject = ""
                sender = ""
                for header in headers:
                    if header["name"] == "Subject":
                        subject = header["value"]
                    if header["name"] == "From":
                        sender = header["value"]
                list_view.insert(END, f"From: {sender}\nSubject: {subject}\n")
        else:
            list_view.insert(END, "No unread messages\n")
        
        # Removing the loading label
        loading_label.destroy()
    
    # Starting the thread
    threading.Thread(target=load_emails).start()

# Creating the show compose function
def show_compose():
    # Creating a new window
    compose_window = Toplevel(window)
    compose_window.title("Compose")
    compose_window.geometry("350x700")
    compose_window.configure(bg="#f0f0f0")
    
    # Creating the sender label
    sender_label = Label(compose_window, text="Sender:", bg="#f0f0f0", font=("Arial", 12))
    sender_label.pack(pady=10)
    
    # Creating the recipient label and entry
    recipient_label = Label(compose_window, text="Recipient:", bg="#f0f0f0", font=("Arial", 12))
    recipient_label.pack(pady=10)
    recipient_entry = Entry(compose_window, width=30, bd=2)
    recipient_entry.pack()

    # Creating the subject label and entry
    subject_label = Label(compose_window, text="Subject:", bg="#f0f0f0", font=("Arial", 12))
    subject_label.pack(pady=10)
    subject_entry = Entry(compose_window, width=30, bd=2)
    subject_entry.pack()

    # Creating the message label and text area
    message_label = Label(compose_window, text="Message:", bg="#f0f0f0", font=("Arial", 12))
    message_label.pack(pady=10)
    message_text = Text(compose_window, width=40, height=20, bd=2)
    message_text.pack()

    # Creating the send email function
    def send_email():
        # Getting the sender, recipient, subject and message from the GUI
        sender = "user_email" # Use the user_email variable from the show_inbox function
        recipient = recipient_entry.get()
        subject = subject_entry.get()
        message = message_text.get(1.0, END)
        
        # Checking if the recipient and message are not empty
        if recipient and message:
            # Creating the email message object
            email = EmailMessage()
            email["From"] = sender
            email["To"] = recipient
            email["Subject"] = subject
            email.set_content(message)
            
            # Sending the email using the Gmail API
            try:
                SERVICE_STATUS_HANDLE.users().messages().send(userId="me", body=email).execute()
                messagebox.showinfo("Success", "Email sent successfully")
            except:
                messagebox.showerror("Error", "Email could not be sent")
        else:
            messagebox.showwarning("Warning", "Please enter the recipient and message")

    # Creating the send button
    send_button = Button(compose_window, text="Send Email", command=send_email, bg="#00ff00", font=("Arial",12))
    # Displaying the sender email address on the top of the window
    sender_email = Label(window, text="user_email", bg="#f0f0f0", font=("Arial", 12))
    canvas.create_window(175, 185, window=sender_email)

    # Creating the voice button for recipient
    voice_button_recipient = Button(compose_window, text="Voice Input for Recipient", command=lambda: recipient_entry.insert(END, recognize_speech()), bg="#ffff00", font=("Arial", 12))
    voice_button_recipient.pack(pady=10)
        
    # Creating the voice button for subject
    voice_button_subject = Button(compose_window, text="Voice Input for Subject", command=lambda: subject_entry.insert(END, recognize_speech()), bg="#ffff00", font=("Arial", 12))
    voice_button_subject.pack(pady=10)
        
    # Creating the voice button for message
    voice_button_message = Button(compose_window, text="Voice Input for Message", command=lambda: message_text.insert(END, recognize_speech()), bg="#ffff00", font=("Arial", 12))
    voice_button_message.pack(pady=10)
window.mainloop()

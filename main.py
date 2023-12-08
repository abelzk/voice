# Importing the modules
import smtplib
import speech_recognition as sr
import pyttsx3
from email.message import EmailMessage
from tkinter import *
from tkinter import messagebox

# Creating the GUI window
window = Tk()
window.title("Voice Based Email Sender")
window.geometry("350x700")
window.configure(bg="#f0f0f0")

# Creating the sender label
sender_label = Label(window, text="Sender: abelzeki24@gmail.com", bg="#f0f0f0", font=("Arial", 12))
sender_label.pack(pady=10)

# Creating the recipient label and entry
recipient_label = Label(window, text="Recipient:", bg="#f0f0f0", font=("Arial", 12))
recipient_label.pack(pady=10)
recipient_entry = Entry(window, width=30, bd=2)
recipient_entry.pack()

# Creating the subject label and entry
subject_label = Label(window, text="Subject:", bg="#f0f0f0", font=("Arial", 12))
subject_label.pack(pady=10)
subject_entry = Entry(window, width=30, bd=2)
subject_entry.pack()

# Creating the message label and text area
message_label = Label(window, text="Message:", bg="#f0f0f0", font=("Arial", 12))
message_label.pack(pady=10)
message_text = Text(window, width=40, height=20, bd=2)
message_text.pack()

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

# Creating the send email function
def send_email():
    # Getting the sender, recipient, subject and message from the GUI
    sender = "abelzeki24@gmail.com"
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
        
        # Sending the email using SMTP
        try:
            with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
                smtp.login(sender, "password") # Replace password with the actual password
                smtp.send_message(email)
                messagebox.showinfo("Success", "Email sent successfully")
        except:
            messagebox.showerror("Error", "Email could not be sent")
    else:
        messagebox.showwarning("Warning", "Please enter the recipient and message")

# Creating the send button
send_button = Button(window, text="Send Email", command=send_email, bg="#05f", font=("Arial", 12))
send_button.pack(pady=10)

# Creating the voice button
voice_button = Button(window, text="Voice Input", command=lambda: message_text.insert(END, recognize_speech()), bg="#ffff00", font=("Arial", 12))
voice_button.pack(pady=10)

# Running the main loop
window.mainloop()

from pynput import keyboard
import speech_recognition as sr
import pyautogui
from docx import Document
from datetime import datetime

text_entries = []
image_paths = []

def capture_screenshot():
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    screenshot_path = f"screenshot_{timestamp}.png"
    pyautogui.screenshot(screenshot_path)
    print(f"Screenshot saved as {screenshot_path}")
    return screenshot_path

def transcribe_speech():
    recognizer = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening for speech...")
        audio = recognizer.listen(source)
        try:
            text = recognizer.recognize_google(audio)
            print(f"Transcribed Text: {text}")
            return text
        except sr.UnknownValueError:
            print("Google Speech Recognition could not understand the audio")
            return ""
        except sr.RequestError as e:
            print(f"Could not request results from Google Speech Recognition service; {e}")
            return ""

def create_document():
    doc = Document()
    doc.add_heading("SOP Document", level=1)
    for text in text_entries:
        doc.add_paragraph(text)
    for image_path in image_paths:
        doc.add_picture(image_path)
    doc_filename = f"SOP_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
    doc.save(doc_filename)
    print(f"Document saved as {doc_filename}")

def on_press(key):
    try:
        if key == keyboard.HotKey.parse('<ctrl>+<shift>+f'):
            text = transcribe_speech()
            if text:
                text_entries.append(text)
        elif key == keyboard.HotKey.parse('<ctrl>+<shift>+p'):
            image_path = capture_screenshot()
            image_paths.append(image_path)
        elif key == keyboard.HotKey.parse('<ctrl>+<shift>+q'):
            create_document()
            return False  # Stop listener
    except AttributeError:
        pass

with keyboard.Listener(on_press=on_press) as listener:
    listener.join()

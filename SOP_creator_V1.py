from pynput import keyboard, mouse
import speech_recognition as sr
import pyautogui
from docx import Document
from datetime import datetime
from PIL import Image, ImageDraw

text_entries = []
image_paths = []

def capture_screenshot_with_highlight(x, y):
    # Capture a screenshot
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    screenshot_path = f"screenshot_{timestamp}.png"
    screenshot = pyautogui.screenshot()

    # Convert screenshot to a PIL Image to draw on it
    screenshot = screenshot.convert("RGB")
    draw = ImageDraw.Draw(screenshot)

    # Define the size of the rectangle around the cursor
    rect_size = 50
    left = x - rect_size // 2
    top = y - rect_size // 2
    right = x + rect_size // 2
    bottom = y + rect_size // 2

    # Draw a rectangle around the cursor position
    draw.rectangle([left, top, right, bottom], outline="red", width=3)

    # Save the screenshot with the highlighted cursor area
    highlighted_path = f"highlighted_{timestamp}.png"
    screenshot.save(highlighted_path)
    print(f"Screenshot with highlight saved as {highlighted_path}")
    return highlighted_path

def transcribe_speech():
    recognizer = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening for speech...")
        audio = recognizer.listen(source)
        try:
            # Recognize speech using Google Web Speech API
            text = recognizer.recognize_google(audio)
            print(f"Transcribed Text: {text}")
            return text
        except sr.UnknownValueError:
            print("Google Speech Recognition could not understand the audio")
            return ""
        except sr.RequestError as e:
            print(f"Could not request results from Google Speech Recognition service; {e}")
            return ""

def create_document(text_entries, image_paths):
    doc = Document()
    doc.add_heading("SOP Document", level=1)
    
    for text in text_entries:
        doc.add_paragraph(text)
    
    for image_path in image_paths:
        doc.add_picture(image_path)
    
    doc_filename = f"SOP_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
    doc.save(doc_filename)
    print(f"Document saved as {doc_filename}")

def on_click(x, y, button, pressed):
    if pressed:
        print(f"Mouse clicked at ({x}, {y})")
        image_path = capture_screenshot_with_highlight(x, y)
        image_paths.append(image_path)
        return False  # Stop listener after the first click to avoid multiple captures

def on_press(key):
    try:
        if key == keyboard.HotKey.parse('<ctrl>+<shift>+f'):
            text = transcribe_speech()
            if text:
                text_entries.append(text)
        elif key == keyboard.HotKey.parse('<ctrl>+<shift>+p'):
            print("Click anywhere on the screen to capture a screenshot with a highlight.")
            with mouse.Listener(on_click=on_click) as listener:
                listener.join()
        elif key == keyboard.HotKey.parse('<ctrl>+<shift>+q'):
            create_document(text_entries, image_paths)
            return False  # Stop listener
    except AttributeError:
        pass

if __name__ == "__main__":
    print("Press 'ctrl+shift+f' to start/stop speech recognition.")
    print("Press 'ctrl+shift+p' to capture a screenshot.")
    print("Press 'ctrl+shift+q' to create the document and exit.")
    
    with keyboard.Listener(on_press=on_press) as listener:
        listener.join()

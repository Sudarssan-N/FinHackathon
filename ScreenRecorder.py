import tkinter as tk
import pyautogui
import numpy as np
import cv2
from datetime import datetime
import threading

recording = False

# Function to start screen recording
def start_recording():
    global recording
    recording = True
    
    # Get the screen size
    screen_size = pyautogui.size()
    
    # Define the codec and create a VideoWriter object
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    file_name = f"screen_recording_{timestamp}.avi"
    fourcc = cv2.VideoWriter_fourcc(*"XVID")
    out = cv2.VideoWriter(file_name, fourcc, 20.0, screen_size)
    
    # Start recording the screen
    while recording:
        img = pyautogui.screenshot()
        frame = np.array(img)
        frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
        out.write(frame)
    
    # Release the VideoWriter object once recording is stopped
    out.release()
    print(f"Screen recording saved as {file_name}")

# Function to stop screen recording
def stop_recording():
    global recording
    recording = False

# Function to handle screen recording in a separate thread
def record_screen():
    record_thread = threading.Thread(target=start_recording)
    record_thread.start()

# GUI setup
root = tk.Tk()
root.title("Screen Recorder")

start_button = tk.Button(root, text="Start Recording", command=record_screen)
start_button.pack(pady=10)

stop_button = tk.Button(root, text="Stop Recording", command=stop_recording)
stop_button.pack(pady=10)

root.mainloop()

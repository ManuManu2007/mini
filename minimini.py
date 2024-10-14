#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import speech_recognition as sr
import pyautogui
import pyttsx3
import comtypes.client
import os

# Initialize the recognizer
recognizer = sr.Recognizer()

# Initialize pyttsx3 for text-to-speech
engine = pyttsx3.init()

# Initialize PowerPoint
def open_ppt(ppt_path):
    if not os.path.exists(ppt_path):
        print(f"Error: File not found at {ppt_path}")
        return None

    powerpoint = comtypes.client.CreateObject("PowerPoint.Application")
    powerpoint.Visible = 1
    presentation = powerpoint.Presentations.Open(ppt_path)
    try:
        presentation.SlideShowSettings.Run()  # Try to start slideshow mode
    except Exception as e:
        print(f"Failed to start slideshow: {e}")
        speak_text("Failed to start slideshow mode.")
    return presentation

def speak_text(text):
    """Converts the given text to speech."""
    engine.say(text)
    engine.runAndWait()

def recognize_speech():
    """Listens for a command and returns it as text."""
    with sr.Microphone() as source:
        print("Listening for commands...")
        recognizer.adjust_for_ambient_noise(source, duration=1)
        audio = recognizer.listen(source)

        try:
            command = recognizer.recognize_google(audio).lower()
            print(f"You said: {command}")
            return command
        except sr.UnknownValueError:
            print("Sorry, I did not understand that.")
            speak_text("Sorry, I did not understand that.")
            return None
        except sr.RequestError:
            print("Network error. Please check your internet connection.")
            speak_text("Network error. Please check your internet connection.")
            return None

def control_ppt(command, presentation):
    try:
        if presentation.SlideShowWindow:  # Check if slideshow is running
            slide_show = presentation.SlideShowWindow.View  # Access the slideshow view

            if "next" in command or "down" in command:
                pyautogui.press('right')  # Move to next slide
                speak_text("Next slide")
            elif "previous" in command or "up" in command:
                pyautogui.press('left')  # Move to previous slide
                speak_text("Previous slide")
            elif "first" in command:
                pyautogui.press('home')  # Go to the first slide
                speak_text("First slide")
            elif "last" in command:
                pyautogui.press('end')  # Go to the last slide
                speak_text("Last slide")
            elif "full screen" in command:
                pyautogui.hotkey('shift', 'f5')  # Start slideshow from the current slide
                speak_text("Showing the current slide in full screen")
            elif "back" in command:
                pyautogui.press('esc')  # Exit full-screen mode
                speak_text("Returning to normal mode")
            elif "exit" in command or "close" in command:
                speak_text("Exiting PowerPoint presentation.")
                return "exit"  # Return "exit" to stop the loop
            elif "new slide" in command:  # Create a new blank slide
                try:
                    new_slide_index = presentation.Slides.Count + 1
                    presentation.Slides.Add(new_slide_index, 6)  # Add blank slide (layout 6)
                    speak_text("New blank slide created.")
                    print(f"New blank slide created at index: {new_slide_index}")
                except Exception as e:
                    speak_text(f"Error while creating a new slide: {e}")
                    print(f"Error: {e}")
            elif "delete slide" in command:  # Delete current slide
                try:
                    current_slide_index = slide_show.Slide.SlideIndex
                    presentation.Slides[current_slide_index].Delete()
                    speak_text("Current slide deleted.")
                    print(f"Slide {current_slide_index} deleted.")
                except Exception as e:
                    speak_text(f"Error while deleting slide: {e}")
                    print(f"Error: {e}")
        else:
            speak_text("No slideshow is currently running. Please start the slideshow.")
            print("No slideshow is currently running.")
    except Exception as e:
        speak_text("There was an error with PowerPoint.")
        print(f"Error: {e}")

def main(ppt_path):
    # Open PowerPoint
    presentation = open_ppt(ppt_path)
    if presentation:
        print("PowerPoint opened.")

    try:
        while True:
            command = recognize_speech()
            if command:
                # If the command is "exit", break the loop and stop the program
                if "exit" in command or "close" in command:
                    speak_text("Exiting program.")
                    break  # Break the loop when "exit" command is recognized
                # Process other PowerPoint commands
                elif any(word in command for word in ["next", "previous", "first", "last", "full screen", "back", "new slide", "delete slide"]):
                    control_ppt(command, presentation)
                else:
                    speak_text("Sorry, I can't perform that command.")
                    print(f"Invalid command: {command}")
    except KeyboardInterrupt:
        print("\nProgram interrupted by user. Exiting gracefully.")
        speak_text("Program interrupted. Exiting now.")
    except Exception as e:
        print(f"An error occurred: {e}")
    finally:
        if presentation:
            try:
                presentation.Close()  # Attempt to close the presentation if it's still open
                print("Presentation closed successfully.")
            except Exception as e:
                print(f"Failed to close the presentation: {e}")
        engine.stop()

if __name__ == "__main__":
    # Path to your PowerPoint file
    ppt_file_path = r"C:\Users\manoj\OneDrive\Documents\IMPLEMENTION OF CLASSIFICATION MODELS.pptx"
    main(ppt_file_path)


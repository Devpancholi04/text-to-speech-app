import win32com.client as wincl
import tkinter as tk
from tkinter import *
import os
import sys

#https://stackoverflow.com/questions/31836104/pyinstaller-and-onefile-how-to-include-an-image-in-the-exe-file
def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS2
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

Logo = resource_path("Logo.png")


def male_voice():
    speaker = wincl.Dispatch("SAPI.SpVoice")
    voices = speaker.GetVoices()
    speaker.Rate = -2
    speaker.Voice = voices.Item(0)
    text = entry.get("1.0", END)
    speaker.speak(text)


def female_voice():
    speaker = wincl.Dispatch("SAPI.SpVoice")
    voices = speaker.GetVoices()
    speaker.Rate = -2
    text = entry.get("1.0", END)
    speaker.Voice = voices.Item(1)
    speaker.speak(text)

def run_function():
    if selected_option.get() == "MALE VOICE":
        male_voice()
    elif selected_option.get() == "FEMALE VOICE":
        female_voice()
    else:
        return "select a vaild option"
    
def exit_application():
    window.destroy()

window = tk.Tk()
window.title("TEXT TO SPEECH")
window.geometry("700x500")

background = PhotoImage(file = resource_path("assets\\back2.png"))
#background = PhotoImage(file = "D:\\programming\\project\\text to speech\\dist\\assets")

background_label = tk.Label(window, image = background)
background_label.place(relwidth = 1, relheight = 1.05)

label = tk.Label(window, text = "ENTER THE TEXT HERE ", font = ("",15)) 
label.pack(pady = 20)

entry = tk.Text(window,width=40 ,height=5, font=("",20))
entry.pack(ipadx=5, ipady=5)

label = tk.Label(window, text = "choose voice type ", font = ("",15))
label.pack(pady=15)

option =['SELECT A OPTION',"MALE VOICE", "FEMALE VOICE"] 

selected_option = tk.StringVar(window)
selected_option.set(option[0])

option_menu = tk.OptionMenu(window,selected_option, *option)
option_menu.pack(pady=15)
option_menu.config(font=("", 10))

run = tk.Button(window, text = "RUN", command = run_function, font = ("",10), width = 20)
run.pack(pady=5)

exit = tk.Button(window, text = "EXIT", command= exit_application, font = ("",10), width = 20)
exit.pack()

window.mainloop()

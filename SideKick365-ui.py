# Simple tkinter UI for SideKick365, a tool to generate PowerPoint presentations from Word documents using GPT-4 and DALL-E
# Steve Goodman 2023/09/27

import argparse
import os
import sys
import sidekick365
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox

# create a window
window = Tk()
window.title("SideKick365")
window.geometry('400x300')

# set the window to always on top
window.attributes("-topmost", True)

# create a label in large text and center it
lbl = Label(window, text="SideKick365", font=("Segeo UI SemiBold", 14), anchor="center")
lbl.grid(column=0, row=0, columnspan=2)

# create a text box for the word file
wordfile = StringVar()
wordfile.set("")

# create an upload button for the word file
def uploadwordfile():
    wordfile.set(filedialog.askopenfilename(initialdir = ".",title = "Select word file",filetypes = (("Word files","*.docx"),("all files","*.*"))))
    return

uploadwordfile_button = Button(window, text="Select a document to transform...", justify="center", command=uploadwordfile)
uploadwordfile_button.grid(column=0, row=1, columnspan=2)

# create a text box for the custom phrase
customphrase = StringVar()
customphrase.set("")
customphrase_label = Label(window, text="Let me know how you want it customized..", justify="center", wraplength=300)
customphrase_label.grid(column=0, row=2, columnspan=2, rowspan=2)

customphrase_textbox = Entry(window,width=30,textvariable=customphrase, justify="center")
customphrase_textbox.grid(column=0, row=5, columnspan=2, rowspan=2)

# Create a button to generate the powerpoint
def generatepowerpoint():
    # check the wordfile exists
    if os.path.isfile(wordfile.get()) == False:
        messagebox.showinfo('Error', 'The wordfile does not exist')
        return
    # set the powerpointfile to the wordfile with a .pptx extension and remove the docx extension
    powerpointfile = wordfile.get().replace(".docx","") + ".pptx"
    # check the powerpointfile does not exist
    
    if os.path.isfile(powerpointfile) == True:
        messagebox.showinfo('Error', 'The PowerPoint file already exists')
        return
    # create a label that says "Generating PowerPoint..." over the height and width of the window
    generatingpowerpoint_label = Label(window, text="Generating PowerPoint using GPT-4 & DALL-E...", font=("Segeo UI SemiBold", 16), justify="center", wraplength=300)
    generatingpowerpoint_label.grid(column=0, row=3, columnspan=2, rowspan=5)
    slidecount=sidekick365.GeneratePowerPointFromWord(wordfile.get(), powerpointfile, customphrase.get())
    generatingpowerpoint_label = Label(window, text="Finishing up in PowerPoint...", font=("Segeo UI SemiBold", 16), justify="center", wraplength=300)
    sidekick365.OpenPowerPointAndApplyDesigner(powerpointfile,slidecount)
    # destroy the label after 5 seconds
    generatingpowerpoint_label.after(5000, generatingpowerpoint_label.destroy)
    return

generatepowerpoint_button = Button(window, text="Generate an AI PowerPoint", command=generatepowerpoint)
generatepowerpoint_button.grid(column=0, row=7, columnspan=2, rowspan=2)


# start the window
window.mainloop()

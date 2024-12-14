import tkinter
from tkinter import filedialog, Label, Button, Entry, IntVar, StringVar, messagebox
from path import Path
from PyPDF4.pdf import PdfFileWriter, PdfFileReader
import pyttsx3
from speech_recognition import Recognizer, AudioFile, RequestError, UnknownValueError
from pydub import AudioSegment
import os

from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

# Declaring global variables
global end_pgNo, start_pgNo, pdfPath

# Function to read text from the PDF and convert it to speech
def read():
    path = filedialog.askopenfilename(title="Select PDF File", filetypes=[("PDF Files", "*.pdf")])  # Get the path of the PDF
    if not path:
        return

    pdfLoc = open(path, 'rb')
    pdfreader = PdfFileReader(pdfLoc)

    start = start_pgNo.get()  # Get start page
    end = end_pgNo.get()  # Get end page

    # Initialize the speaker
    speaker = pyttsx3.init()

    # Ensure page numbers are within the bounds
    num_pages = pdfreader.numPages
    if start < 1 or end > num_pages or start > end:
        messagebox.showerror("Error!", "Please enter valid page numbers.")
        return

    # Read and speak the text for each page in the specified range
    for i in range(start - 1, end):  # Adjust for zero-based index
        page = pdfreader.getPage(i)
        txt = page.extractText()

        if txt.strip():  # Check if there's text to read
            print(f"Extracted text from page {i + 1}: {txt}")  # Debugging output
            speaker.say(txt)
            speaker.runAndWait()  # Wait for the speech to finish
        else:
            print(f"No text found on page {i + 1}.")
    
    pdfLoc.close()  # Close the PDF file after reading

# Function to create GUI for PDF to audio conversion
def pdf_to_audio():
    global start_pgNo, end_pgNo
    wn1 = tkinter.Tk()
    wn1.title("PDF to Audio converter")
    wn1.geometry('500x400')
    wn1.config(bg='snow3')
    start_pgNo = IntVar(wn1)
    end_pgNo = IntVar(wn1)

    Label(wn1, text='PDF to Audio Converter', fg='black', font=('Courier', 15)).place(x=60, y=10)
    Label(wn1, text='Enter start and end page:', anchor="e", justify="left").place(x=20, y=90)
    Label(wn1, text='Start Page No.:').place(x=100, y=140)
    startPg = Entry(wn1, width=20, textvariable=start_pgNo)
    startPg.place(x=100, y=170)
    Label(wn1, text='End Page No.:').place(x=250, y=140)
    endPg = Entry(wn1, width=20, textvariable=end_pgNo)
    endPg.place(x=250, y=170)
    Button(wn1, text="Read PDF", bg='ivory3', font=('Courier', 13), command=read).place(x=200, y=260)
    wn1.mainloop()

# Function to write text into a PDF using reportlab
def write_text(filename, text):
    c = canvas.Canvas(filename, pagesize=letter)  # Create a new PDF canvas
    c.setFont("Helvetica", 12)  # Set the font and size
    width, height = letter  # Get the page size
    
    # Write text to PDF, adjusting the position
    lines = text.splitlines()
    y = height - 40  # Start near the top of the page
    for line in lines:
        # Split long lines into multiple lines
        words = line.split(" ")
        current_line = ""
        for word in words:
            # Check if adding this word would exceed the width of the page
            test_line = current_line + " " + word if current_line else word
            text_width = c.stringWidth(test_line, "Helvetica", 12)
            if text_width < width - 40:  # Leave some margin
                current_line = test_line
            else:
                c.drawString(40, y, current_line)  # Draw the current line
                y -= 15  # Move down for the next line
                current_line = word  # Start a new line with the current word
            
            # Check for page overflow
            if y < 40:
                c.showPage()
                c.setFont("Helvetica", 12)  # Reset font on new page
                y = height - 40  # Reset y position for new page

        # Draw any remaining text in the current line
        if current_line:
            c.drawString(40, y, current_line)
            y -= 15  # Move down for the next line
            
    c.save()  # Save the PDF file

# Function to convert audio into text and save to PDF
def convert():

    path = filedialog.askopenfilename(title="Select Audio File", filetypes=[("Audio Files", "*.wav;*.mp3")])  # Get the audio file path
    if not path:
        return

    # Ask user to select folder to save the PDF
    pdf_loc = filedialog.asksaveasfilename(defaultextension='.pdf', filetypes=[("PDF Files", "*.pdf")])
    if not pdf_loc:
        messagebox.showerror('Error!', 'Please choose a location to save the PDF.')
        return

    audioFileName = os.path.basename(path).split('.')[0]
    audioFileExt = os.path.splitext(path)[1][1:]

    if audioFileExt not in ['wav', 'mp3']:
        messagebox.showerror('Error!', 'Audio format should be "wav" or "mp3".')
        return

    # Convert mp3 to wav if necessary
    if audioFileExt == 'mp3':
        audio_file = AudioSegment.from_file(path, format='mp3')
        source_file = f'{audioFileName}.wav'
        audio_file.export(source_file, format='wav')
    else:
        source_file = path

    recog = Recognizer()
    try:
        with AudioFile(source_file) as source:
            recog.pause_threshold = 5
            audio_duration = int(source.DURATION)
            
            if audio_duration <= 0:
                messagebox.showerror("Error", "The audio file seems to be empty.")
                return

            # Break down the audio into chunks if it's more than 60 seconds
            chunk_size = 60  # seconds
            text = ""
            for i in range(0, audio_duration, chunk_size):
                audio_chunk = recog.record(source, duration=min(chunk_size, audio_duration - i))
                try:
                    chunk_text = recog.recognize_google(audio_chunk)
                    text += chunk_text + " "
                except UnknownValueError:
                    messagebox.showwarning("Error", "Google could not understand the audio in some parts.")
                    text += "[Unintelligible] "
                except RequestError as e:
                    messagebox.showerror('Error!', f"Could not request results from Google API: {str(e)}")
                    return

            print("Recognized Text:")
            print(text)
            
            # Save the recognized text to PDF
            write_text(pdf_loc, text)
            messagebox.showinfo('Success!', f'Audio converted to PDF: {os.path.basename(pdf_loc)}')

    except Exception as e:
        messagebox.showerror('Error!', f'An error occurred during audio processing: {str(e)}')

# Function to create GUI for audio to PDF conversion
def audio_to_pdf():
    global pdfPath
    wn2 = tkinter.Tk()
    wn2.title("Audio to PDF converter")
    wn2.geometry('500x400')
    wn2.config(bg='snow3')
    pdfPath = StringVar(wn2)

    Label(wn2, text='Audio to PDF Converter', fg='black', font=('Courier', 15)).place(x=60, y=10)
    Label(wn2, text='Choose the audio file location (.wav or .mp3):').place(x=20, y=130)
    Button(wn2, text='Choose', bg='ivory3', font=('Courier', 13), command=convert).place(x=50, y=170)
    wn2.mainloop()

# Main window with two options
wn = tkinter.Tk()
wn.title("PDF to Audio and Audio to PDF converter")
wn.geometry('700x300')
wn.config(bg='LightBlue1')
Label(wn, text='PDF to Audio and Audio to PDF Converter', fg='black', font=('Courier', 15)).place(x=40, y=10)
Button(wn, text="Convert PDF to Audio", bg='ivory3', font=('Courier', 15), command=pdf_to_audio).place(x=230, y=80)
Button(wn, text="Convert Audio to PDF", bg='ivory3', font=('Courier', 15), command=audio_to_pdf).place(x=230, y=150)
wn.mainloop()

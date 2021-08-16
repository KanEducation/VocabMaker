import pandas as pd
import random
import os
from fpdf import FPDF
from tkinter import *
import tkinter.font



# Function that draws line for the first page
def lines(self):
    self.set_fill_color(0, 0, 0) # color for outer rectangle
    self.rect(1.0, 1.0, 208.0,296.0,'DF')
    self.set_fill_color(255, 255, 255) # color for inner rectangle
    self.rect(1.5, 1.5, 207.0,295.0,'FD')

def imagex(self, path):
        self.set_xy(4.0,3.0)
        self.image(path, link='', type='', w=10, h=15)

#Creates word qs and numbers them
def word_creator(self, selected_word):
    q_number = 1
    for word in selected_word:
        current_word = str(q_number) + '. ' + word
        self.set_font('Arial', 'B', 12)
        self.cell(210, 20, txt = current_word, ln = 2, align = 'L')
        q_number = q_number + 1


def answer_creator(self, selected_answer):
    q_number = 1
    for answer in selected_answer:
        current_answer = str(q_number) + '. ' + answer
        self.set_font('Arial', 'B', 12)
        self.cell(210, 20, txt = current_answer, ln = 2, align = 'L')
        q_number = q_number + 1
    



#Check if there already exists a vocab quiz file in the current directory
def alert_popup(title, message, path):
    """Generate a pop-up window for special messages."""
    root = Tk()
    root.title(title)
    font = tkinter.font.Font(family="Aria", size=11)
    w = 400     # popup window width
    h = 200     # popup window height
    sw = root.winfo_screenwidth()
    sh = root.winfo_screenheight()
    x = (sw - w)/2
    y = (sh - h)/2
    root.geometry('%dx%d+%d+%d' % (w, h, x, y))
    m = message
    m += '\n'
    m += path
    w = Label(root, text=m, width=120, height=10, font = font)
    w.pack()
    b = Button(root, text="OK", command=root.destroy, width=10, font=font)
    b.pack()
    mainloop()

fname = os.path.join(os.getcwd(),'vocab_quiz.pdf')
if (os.path.isfile(fname)):
    alert_popup("Error", "There already exists a vocab quiz file. \n Please remove it first!", os.getcwd())
else:
    df = pd.read_excel('Vocab.xlsx')


    meaning_store = df['Meaning']
    word_store = df['Word']
    select_meaning = []
    select_word = []

    total_word = len(word_store)
    index = [i for i in range (0,len(word_store), 1)]

    # For 0.8 of the entire word list, randomly take in (X)
    # Fix: Grab only 20

    for i in range(0,20,1):
        select_index = random.choice(index)
        select_word.append(word_store[select_index])
        select_meaning.append(meaning_store[select_index])
        index.remove(select_index)
        pdf = FPDF() 
        pdf.add_page()
        pdf.set_font("Arial", size = 15)


    #CREATES THE VOCAB QUIZ SHEET
    lines(pdf)
    imagex(pdf,os.path.join(os.getcwd(),'kan_logo.png'))
    #titles(pdf, 'Vocab Quiz')
    pdf.set_font('Arial', 'B', 15)
    pdf.text(x = 100, y = 10, txt="Vocab Quiz")
    word_creator(pdf, select_word)
    pdf.output("vocab_quiz.pdf")   

    #CREATES THE ANSWER SHEET
    pdf_answer = FPDF() 
    pdf_answer.add_page()
    pdf_answer.set_font('Arial', 'B', 14)
    imagex(pdf_answer,os.path.join(os.getcwd(),'kan_logo.png'))
    pdf_answer.text(x = 100, y = 10, txt="Answer Sheet")
    answer_creator(pdf_answer, select_meaning)
    pdf_answer.output("vocab_quiz_answer.pdf")   




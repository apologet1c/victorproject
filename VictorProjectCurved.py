import docx2txt
import sys
import jellyfish
import os
import os.path
from tkinter import *
from tkinter import messagebox
import win32com.client

#change window height and add printbox
window = Tk()
window.title("Bill Analysis Analyzer")
window.geometry('350x360')

# this section is the actual analyzer so it can be called by the gui/tkinter later
def analyzer(hs,bill):
    if hs.get() == 2:
        hs = "SB"
    elif hs.get() == 3:
        hs = "HJR"
    elif hs.get() == 4:
        hs = "SJR"
    else:
        hs = "HB" #because most bills will probably be house bills--this makes it default to HB

    if int(bill) < 10:
        abill = "0000" + bill
    elif int(bill) < 100:
        abill = "000" + bill
    elif int(bill) < 1000:
        abill = "00" + bill
    elif int(bill) > 1000 and int(bill) < 10000:
        abill = "0" + bill
    else:
        quote = "We're sorry, but that doesn't seem like the correct bill number. Please hang up and try the number again.\n"
        T.insert(END, quote)
    quote = "Now analyzing " + hs.upper() + bill + ".....\n"
    T.insert(END, quote)

    # generating filenames
    filename = "X:/HRO/DocumentStorage/drafting/ba/85_R/" + hs + abill + ".docx"
    ofilename = "X:/HRO/DocumentStorage/drafting/ba/85_R/Original/"+ hs + abill + ".docx"

    # checking if the files exist
    path1 = str(ofilename)
    path2 = str(filename)
    if os.path.isfile(path1) and os.access(path1, os.R_OK):
       asdf = 1 #I don't know how to do this properly
    else:
        quote = "Sorry, but I couldn't find the original version of that bill analysis.\n"
        T.insert(END, quote)
    if os.path.isfile(path2) and os.access(path2, os.R_OK):
       asdf = 1 #I don't know how to do this properly
    else:
        quote = "Sorry, but I couldn't find the edited version of that bill analysis.\n\n"
        T.insert(END, quote)
        
    # converting files to plaintext and removing line breaks
    edited = docx2txt.process(filename)
    editednobr = edited.replace('\r', '').replace('\n', '')
    original = docx2txt.process(ofilename)
    originalnobr = original.replace('\r', '').replace('\n', '')

    # splitting into sections, ORIGINAL +++++++++++++++++++++++++
    # for bills without backgrounds
    if "BACKGROUND:" in originalnobr:
        headingo, rest1o = originalnobr.split("BACKGROUND:")
    else:
        headingo, rest1o = originalnobr.split("DIGEST:")

    # for bills with arguments and notes
    if "SUPPORTERSSAY:" in originalnobr and "NOTES:" in originalnobr:
        backdigesto, rest2o = rest1o.split("SUPPORTERSSAY:")
        argumentso, noteso = rest2o.split("NOTES:")

    # bills with arguments but with no notes
    elif "SUPPORTERSSAY:" in originalnobr and "NOTES:" not in originalnobr:
        backdigesto, argumentso = rest1o.split("SUPPORTERSSAY:")
        noteso = "" #in case notes are added by editors

    # bills with no arguments but with notes
    elif "SUPPORTERSSAY:" not in originalnobr and "NOTES:" in originalnobr:
        backdigesto, noteso = rest1o.split("NOTES:")

    #bills with no arguments and no notes
    else:
        backdigesto = rest1o
        noteso = "" #in case notes are added by editors

    # splitting into sections EDITED +++++++++++++++++++++++++ 
    # CASE 1: for bills without backgrounds
    if "BACKGROUND:" in editednobr:
        heading, rest1 = editednobr.split("BACKGROUND:")
        jscoreheading = jellyfish.jaro_distance(headingo, heading)
        #curve scores--delete this line if the curve doesn't work
        jscoreheading = min(max(-19.5065*jscoreheading**2 + 36.4259*jscoreheading - 15.926, 0), 1)
        jscoreheading = round(jscoreheading, 4)
    else:
        heading, rest1 = editednobr.split("DIGEST:")
        jscoreheading = jellyfish.jaro_distance(headingo, heading)
        #curve scores--delete this line if the curve doesn't work
        jscoreheading = min(max(-19.5065*jscoreheading**2 + 36.4259*jscoreheading - 15.926, 0), 1)
        jscoreheading = round(jscoreheading, 4)
        
    # CASE 2: for bills with arguments and notes
    if "SUPPORTERSSAY:" in editednobr and "NOTES:" in editednobr:
        backdigest, rest2 = rest1.split("SUPPORTERSSAY:")
        arguments, notes = rest2.split("NOTES:")
        jscoredigest = jellyfish.jaro_distance(backdigesto, backdigest)
        jscorearguments = jellyfish.jaro_distance(argumentso, arguments)
        jscorenotes = jellyfish.jaro_distance(noteso, notes)
         #curve scores--delete these lines if the curve doesn't work
        jscoredigest = min(max(-19.5065*jscoredigest**2 + 36.4259*jscoredigest - 15.926, 0), 1)
        jscorearguments = min(max(-19.5065*jscorearguments**2 + 36.4259*jscorearguments - 15.926, 0), 1)
        jscorenotes = min(max(-19.5065*jscorenotes**2 + 36.4259*jscorenotes - 15.926, 0), 1)
        #round
        jscoredigest = round(jscoredigest, 4)
        jscorearguments = round(jscorearguments, 4)
        jscorenotes = round(jscorenotes, 4)
    #calculate weighted average
        jscoreaverage = .1*jscoreheading + .5*jscoredigest + .35*jscorearguments + .05*jscorenotes
        jscoreaverage = round(jscoreaverage, 4)
        
    # CASE 3: bills with arguments but with no notes
    elif "SUPPORTERSSAY:" in editednobr and "NOTES:" not in editednobr:
        backdigest, arguments = rest1.split("SUPPORTERSSAY:")
        jscoredigest = jellyfish.jaro_distance(backdigesto, backdigest)       
        jscorearguments = jellyfish.jaro_distance(argumentso, arguments)
         #curve scores--delete these lines if the curve doesn't work
        jscoredigest = min(max(-19.5065*jscoredigest**2 + 36.4259*jscoredigest - 15.926, 0), 1)
        jscorearguments = min(max(-19.5065*jscorearguments**2 + 36.4259*jscorearguments - 15.926, 0), 1)
        #round
        jscoredigest = round(jscoredigest, 4)
        jscorearguments = round(jscorearguments, 4)
        jscorenotes = "not applicable."
    #calculate weighted average
        jscoreaverage = .13*jscoreheading + .53*jscoredigest + .34*jscorearguments
        jscoreaverage = round(jscoreaverage, 4)
        
    # CASE 4: bills with no arguments but with notes
    elif "SUPPORTERSSAY:" not in editednobr and "NOTES:" in editednobr:
        backdigest, notes = rest1.split("NOTES:")
        jscoredigest = jellyfish.jaro_distance(backdigesto, backdigest)
        jscorenotes = jellyfish.jaro_distance(noteso, notes)
        #curve scores--delete these lines if the curve doesn't work
        jscoredigest = min(max(-19.5065*jscoredigest**2 + 36.4259*jscoredigest - 15.926, 0), 1)
        jscorenotes = min(max(-19.5065*jscorenotes**2 + 36.4259*jscorenotes - 15.926, 0), 1)
        #round
        jscoredigest = round(jscoredigest, 4)
        jscorenotes = round(jscorenotes, 4)
        jscorearguments = "not applicable."
    #calculate weighted average
        jscoreaverage = .13*jscoreheading + .82*jscoredigest + .05*jscorenotes
        jscoreaverage = round(jscoreaverage, 4)
        
    # CASE 5: bills with no arguments and no notes
    else:
        backdigest = rest1
        jscoredigest = jellyfish.jaro_distance(backdigesto, backdigest)
        #curve scores--delete this line if the curve doesn't work
        jscoredigest = min(max(-19.5065*jscoredigest**2 + 36.4259*jscoredigest - 15.926, 0))
        #round
        jscoredigest = round(jscoredigest, 4)
        jscorenotes = "not applicable."
        jscorearguments = "not applicable."
    #calculate weighted average
        jscoreaverage = .2*jscoreheading + .8*jscoredigest
        jscoreaverage = round(jscoreaverage, 4)

    #print to gui 
    s5 = "Your overall score is " + str(jscoreaverage)
    s1 = "\nYour heading score is " + str(jscoreheading)
    s2 = "\nYour background/digest score is " + str(jscoredigest)
    s3 = "\nYour arguments score is " + str(jscorearguments)
    s4 = "\nYour notes score is " + str(jscorenotes) + "\n\n"
    quote = s5+ s1 + s2 + s3 + s4 #because it's only printing every other one if I don't consolidate variables. This whole section could be much cleaner but it works.
    T.insert(END, quote)

    #comparison
    if chk_state.get() == True:
        Application=win32com.client.gencache.EnsureDispatch("Word.Application")
        published = Application.Documents.Open(filename) #for some reason it has to be open to compare
        ori = Application.Documents.Open(ofilename)
        Application.CompareDocuments(ori, published)
        published.Close() #closes the documents after the comparison is done
        ori.Close()

lbl = Label(window, text="Hello! Is this a House or Senate bill we're analyzing?")
lbl.place(relx=.5, rely=.05, anchor="c")

hs = IntVar()
hs.set(1)

rad1 = Radiobutton(window, text='HB', value=1, variable=hs)
rad2 = Radiobutton(window, text='SB', value=2, variable=hs)
rad3 = Radiobutton(window, text='HJR', value=3, variable=hs)
rad4 = Radiobutton(window, text='SJR', value=4, variable=hs)

rad1.place(relx=.2, rely=.125, anchor="c")
rad2.place(relx=.4, rely=.125, anchor="c")
rad3.place(relx=.6, rely=.125, anchor="c")
rad4.place(relx=.8, rely=.125, anchor="c")

lbl = Label(window, text="And what's the bill number?")
lbl.place(relx=.5, rely=.20, anchor="c")

txt = Entry(window, width=10)
txt.place(relx=.5, rely=.2575, anchor="c")
txt.focus()

def clicked(event):
    bill = str(txt.get())
    res = "Thanks, now finding that bill."
    lbl.configure(text=res)
    analyzer(hs, bill)
    #need something here to take the output of analzyer()?
    
window.bind('<Return>', clicked) #allows you to press enter instead of clicking the button

btn = Button(window, text="Run Comparison")
btn.bind('<Button-1>', clicked)
btn.place(relx=.5, rely=.355, anchor="c")

chk_state = BooleanVar()
chk_state.set(False) #defaults to not comparing bills
chk = Checkbutton(window, text='Do you want to open a redline automatically?', var=chk_state)
chk.place(relx=.5, rely=.45, anchor="c")

S = Scrollbar(window)
T = Text(window, height=8, width=40)
T.place(relx=.5, rely=.75, anchor="c")
mainloop()

window.mainloop()

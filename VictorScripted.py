import os
import os.path
import sys
import docx2txt
import jellyfish
import win32com.client
import csv

num = list(range(1, 10)) #change this to change what bill numbers are searched
fivenum = []

session = "85_R" #modify this to change what session is searched

hs = input("Would you like to look for HB, SB, HJR, or SJRs?")

# getting a fivedigit number for the filename
for i in num:
    if i < 10:
        abill = "0000" + str(i)
    elif i < 100:
        abill = "000" + str(i)
    elif i < 1000:
        abill = "00" + str(i)
    elif i > 1000 and i < 10000:
        abill = "0" + str(i)

    fivenum.append(abill)

#now we need to check if the bill exists

validnumbers = []
existonumber = []
for i in fivenum:
    ofilename = "X:/HRO/DocumentStorage/drafting/ba/" + session + "/Original/" + hs + i + ".docx"    
    if os.path.isfile(ofilename) and os.access(ofilename, os.R_OK):
        existonumber.append(i)
       
for i in existonumber: #I'm relying on this because there are some missing originals
    efilename = "X:/HRO/DocumentStorage/drafting/ba/" + session + "/" + hs + i + ".docx"
    if os.path.isfile(efilename) and os.access(efilename, os.R_OK):
        validnumbers.append(i)
        
#now to make the list of filenames
efilenames = []
ofilenames = []
for i in validnumbers:
    efilename = "X:/HRO/DocumentStorage/drafting/ba/" + session + "/" + hs + i + ".docx"
    ofilename = "X:/HRO/DocumentStorage/drafting/ba/" + session + "/Original/" + hs + i + ".docx"
    ofilenames.append(ofilename)
    efilenames.append(efilename)

print("Found bills!")

#get authors and wordcount
authorl = []
wordcounts = []
lastsavel = []

for i in ofilenames:
    word = win32com.client.gencache.EnsureDispatch("Word.Application")
    doc = word.Documents.Open(i)
    author = doc.BuiltInDocumentProperties("Author")
    authorl.append(str(author))
    wcs = doc.BuiltInDocumentProperties("Number of Words")
    wordcounts.append(str(wcs))
    lastsave = doc.BuiltInDocumentProperties("Creation Date")
    lastsavel.append(str(lastsave))
    doc.Close()

print("Authors found!")
print("Wordcounts done!")
print("Creation dates found!")

#now to convert to plain text
editednobr = []
originalnobr = []

for i in efilenames:
    edited = docx2txt.process(i)
    edited = edited.replace('\r', '').replace('\n', '')
    editednobr.append(edited)

for i in ofilenames:
    original = docx2txt.process(i)
    original = original.replace('\r', '').replace('\n', '')
    originalnobr.append(original)

#split into sections -- ORIGINAL
headingol = []
backdigestol = []
argumentsol = []
notesol = []

for i in originalnobr:
    if "BACKGROUND:" in i:
        headingo, rest1o = i.split("BACKGROUND:")
    else:
        headingo, rest1o = i.split("DIGEST:")

    # for bills with arguments and notes
    if "SUPPORTERSSAY:" in i and "NOTES:" in i:
        backdigesto, rest2o = rest1o.split("SUPPORTERSSAY:")
        argumentso, noteso = rest2o.split("NOTES:")

    # bills with arguments but with no notes
    elif "SUPPORTERSSAY:" in i and "NOTES:" not in i:
        backdigesto, argumentso = rest1o.split("SUPPORTERSSAY:")
        noteso = " "

    # bills with no arguments but with notes
    elif "SUPPORTERSSAY:" not in i and "NOTES:" in i:
        backdigesto, noteso = rest1o.split("NOTES:")
        argumentso = " "

    #bills with no arguments and no notes
    else:
        backdigesto = rest1o
        noteso = " "
        argumentso = " "

    headingol.append(headingo)
    backdigestol.append(backdigesto)
    argumentsol.append(argumentso)
    notesol.append(noteso)

#split into sections -- EDITED
headingel = []
backdigestel = []
argumentsel = []
notesel = []

for i in editednobr:
    if "BACKGROUND:" in i:
        headinge, rest1e = i.split("BACKGROUND:")
    else:
        headinge, rest1e = i.split("DIGEST:")

    headingol.append(headingo)
    # for bills with arguments and notes
    if "SUPPORTERSSAY:" in i and "NOTES:" in i:
        backdigeste, rest2e = rest1e.split("SUPPORTERSSAY:")
        argumentse, notese = rest2e.split("NOTES:")

    # bills with arguments but with no notes
    elif "SUPPORTERSSAY:" in i and "NOTES:" not in i:
        backdigeste, argumentse = rest1e.split("SUPPORTERSSAY:")
        notese = "x" #so anything without notes comes out as a 0.0

    # bills with no arguments but with notes
    elif "SUPPORTERSSAY:" not in i and "NOTES:" in i:
        backdigeste, noteso = rest1e.split("NOTES:")
        argumentse = "x" #so anything without notes comes out as a 0.0

    #bills with no arguments and no notes
    else:
        backdigeste = rest1e
        notese = "x" #so anything without notes comes out as a 0.0
        argumentse = "x" #so anything without notes comes out as a 0.0

    headingel.append(headinge)
    backdigestel.append(backdigeste)
    argumentsel.append(argumentse)
    notesel.append(notese)

#actually doing the comparison, making lists
    
jhead = []
for (x,y) in zip(headingel, headingol):
    jscorehead = jellyfish.jaro_distance(x, y)
    jhead.append(jscorehead)

print("Heading scores done!")

jbd = []
for (x,y) in zip(backdigestel, backdigestol):
    jscorebd = jellyfish.jaro_distance(x, y)
    jbd.append(jscorebd)

print("Digest scores done!")

jargs = []
for (x,y) in zip(argumentsel, argumentsol):
    jscoreargs = jellyfish.jaro_distance(x, y)
    jargs.append(jscoreargs)

print("Arguments scores done!")

jnotes = []
for (x,y) in zip(notesel, notesol):
    jscorenotes = jellyfish.jaro_distance(x, y)
    jnotes.append(jscorenotes)

print("Notes scores done!")

hslist = []

for i in validnumbers:
    hslist.append(hs)

#saving everything into a CSV
res = zip(hslist, validnumbers, authorl, lastsavel, wordcounts, jhead, jbd, jargs, jnotes)
csvfile = hs + session + "data.csv"

with open(csvfile, "w") as output:
    writer = csv.writer(output, lineterminator='\n')
    writer.writerows(res)

# -*- coding: utf-8 -*-
"""
Created on Wed Apr  8 13:12:18 2020

@author: StP

Starting from .pdf files from a given folder and a template .docx file as
index, create the song book.
"""

import time
import os
import numpy as np
from PyPDF2 import PdfFileReader, PdfFileWriter
from docx import Document
import comtypes.client
start_time = time.time()


# Folders and Files
PdfFolder = os.getcwd() + '\\pdf'
MusxFolder = os.getcwd() + '\\mus'
IndexTemplate = os.getcwd() + '\\Index_Template.docx'
OutputFile = os.getcwd() + '\\LibrettoCanti.pdf'


# List of .pdf Files
PdfFiles = []
for root, dirs, files in os.walk(PdfFolder):
    for filename in files:
        if filename[-4:]=='.pdf':
            PdfFiles.append(filename[:-4])
            

# List of .musx Files
MusxFiles = []
for root, dirs, files in os.walk(MusxFolder):
    for filename in files:
        if filename[-5:]=='.musx':
            MusxFiles.append(filename[:-5])
#MusxFiles = MusxFiles[1:] # Remove 1st entry (template)


# Prepare sorted list of .pdf files (non-case sensitive alphabetical order with some characters removed or modified)
PdfFiles_X = []; PdfFiles_X = PdfFiles[:]
S_Find = np.array(['à', 'è', 'é', 'ì', 'ò', 'ù', '-', '\''])
S_Repl = np.array(['a', 'e', 'e', 'i', 'o', 'u', ' ', ' ' ])
for xfp in range(0, len(PdfFiles)):
    PdfFiles_X[xfp] = PdfFiles[xfp].lower();
    # Replace special characters with regular counterparts
    for xs in range(0, len(S_Find)):
        PdfFiles_X[xfp] = PdfFiles_X[xfp].replace(S_Find[xs], S_Repl[xs])
    # Remove multiple spaces
    while PdfFiles_X[xfp].find('  ')>=0:
        PdfFiles_X[xfp] = PdfFiles_X[xfp].replace('  ', ' ')
NS = np.argsort(PdfFiles_X)
    

# Sort .pdf files
PdfFiles_S = []
PdfFiles_XS = []
for xn in range(0,len(NS)):
    PdfFiles_S.append(PdfFiles[NS[xn]])
    PdfFiles_XS.append(PdfFiles_X[NS[xn]])

    
# Look for EXCEEDING .pdf Files (.musx counterpart not present)
print(' ')
print('====================')        
print(' ')
print('EXCEEDING PDF FILES:')
for xfp in range(0, len(PdfFiles_S)):
    FilePresent = 0
    for xfm in range(0, len(MusxFiles)):
        if PdfFiles_S[xfp]==MusxFiles[xfm]:
            FilePresent = 1
    if FilePresent==0:
        print('* ' + PdfFiles_S[xfp])
print(' ')
print('====================')        
print(' ')


# Look for MISSING .pdf Files (only .musx file present)
print('MISSING PDF FILES:')
for xfm in range(0, len(MusxFiles)):
    FilePresent = 0
    for xfp in range(0, len(PdfFiles)):
        if MusxFiles[xfm]==PdfFiles[xfp]:
            FilePresent = 1
    if FilePresent==0:
        print('* ' + MusxFiles[xfm])
print(' ')
print('====================')        
print(' ')


# Create Index string
IdxStr = ''
for xfp in range(0, len(PdfFiles_S)):
    IdxStr = IdxStr + PdfFiles_S[xfp] + '\n'
IdxStr = IdxStr[:-1]


# Create .docx temporary file with updated Song Index
document = Document(IndexTemplate)
paragraphs = document.paragraphs
paragraphs[-1].text = IdxStr
document.save(os.getcwd() + '\\Index_Temp.docx')


# Export .docx to .pdf
wdFormatPDF = 17
word = comtypes.client.CreateObject('Word.Application')
doc = word.Documents.Open(os.getcwd() + '\\Index_Temp.docx')
doc.SaveAs(os.getcwd() + '\\Index_Temp.pdf', FileFormat=wdFormatPDF)
doc.Close()
word.Quit()


# Delete temporary .docx file
if os.path.exists(os.getcwd() + '\\Index_Temp.docx'):
    os.remove(os.getcwd() + '\\Index_Temp.docx')


# Import Index.pdf
A=PdfFileReader(os.getcwd() + '\\Index_Temp.pdf')
Np = A.getNumPages()
if Np/2!=np.floor(Np/2):
    Np = Np+1
    AddBlankPageAfterIndex = 1
else:
    AddBlankPageAfterIndex = 0
    

# Calculate starting pages
NPageStart = []; NPageStart.append(Np+1)  # 1st file (except for index.pdf) starting page
AddBlankPage = [] # Init Blank Page array. When 1, an extra blank page must be added
for xfp in range(0, len(PdfFiles_S)):
    A=PdfFileReader(os.getcwd() + '\\pdf\\' +  PdfFiles_S[xfp] +'.pdf')
    Np = A.getNumPages()
    if NPageStart[-1]/2!=np.floor(NPageStart[-1]/2) and Np>1:
        # Pdf file with more than 1 page starting from odd page. Add a blank page to make it start from an even one.
        NPageStart.append(NPageStart[-1]+Np+1)
        AddBlankPage.append(1)
    else:
        NPageStart.append(NPageStart[-1]+Np)
        AddBlankPage.append(0)
AddBlankPage.append(0)


# Output .pdf file. Add index
output = PdfFileWriter()
f = os.getcwd() + '\\Index_Temp.pdf'
pdfFile = PdfFileReader(f)
for p in range(pdfFile.getNumPages()):
    output.addPage(pdfFile.getPage(p))
output.addBookmark('INDICE', 2)
if AddBlankPageAfterIndex==1:
    output.addBlankPage()

    
# Output .pdf file. Add other files    
for xfp in range(0, len(PdfFiles_S)):
    if AddBlankPage[xfp]==1:
        output.addBlankPage()
    f = os.getcwd() + '\\pdf\\' +  PdfFiles_S[xfp] + '.pdf'
    pdfFile = PdfFileReader(f)
    for p in range(pdfFile.getNumPages()):
        output.addPage(pdfFile.getPage(p))


# Add Bookmarks
CurrInit = '-'
for xfp in range(0,len(PdfFiles_XS)):
    if PdfFiles_XS[xfp][0]!=CurrInit:
        CurrInit = PdfFiles_XS[xfp][0]
        parent = output.addBookmark(CurrInit.upper(), NPageStart[xfp]-1+AddBlankPage[xfp])
    output.addBookmark(PdfFiles_S[xfp], NPageStart[xfp]-1+AddBlankPage[xfp], parent)


# Collapse Bookmarks
output.addJS('function closeBookmarks(bm)' + \
             '{var i; if (bm.children !== null)' + \
             '{bm.open = false;' + \
             'for (i = 0; i < bm.children.length; i += 1) ' + \
             '{closeBookmarks(bm.children[i]);}}};' + \
             'closeBookmarks(this.bookmarkRoot);' \
             )


# Create output .pdf file
with open(OutputFile, 'wb') as outputStream:
    output.write(outputStream)


# Delete temporary Index.pdf file
if os.path.exists(os.getcwd() + '\\Index_Temp.pdf'):
    os.remove(os.getcwd() + '\\Index_Temp.pdf')
    
    
# Speed report
stop_time = time.time()
DT = stop_time-start_time
print(f'Output PDF File Created. Elapsed Time: {DT:.3f} s')

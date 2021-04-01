# Code Louisville Python project spring 2021

## GOAL

Create python exe program that can run on windows and extract text from a selected document passed via the first command argument.
This program will extract a unique word list from any .doc, .docx or pdf file and save a list of unique list of words to a ??_words.txt file where ?? is the name of the original document.

I have included the test.doc,test.docx and test.pdf for testing


## Required to build:

pip install docx2txt

pip install PyMuPDF

pip install pywin32

pip install docx2txt

pip install pyinstaller  -> Optional - creates EXE for windows


It is designed to run in silent mode, on windows, driven be a C++ program that executes the EXE with the arguement of the file to extract a unique word list.  Once the word list is created the C++ program builds a table cross-referrencing words and docuements, to be able to find all documents with a given combination of words

# How to run this program

You have 2 options:

1 - On Mac or windows, Load into python and run, which will allow you enter the file name to parse, testFile.pdf

2 - Create a windows .exe

    in the terminal type:    pyinstaller docs2txt.py

    run docs2Txt.exe from the dist folder


# Class requirements met
1) At least 3 functions called -> docs2Txt has 4 functions
2) Analyze text and display information -> docs2Txt accepts, input from user, input from document source 
    and generates list of unique words
3) Analyze text and display information  -> docs2Txt scraps the words from 3 different types of documents and outputs a 
    list of unique words in each
4) Implement a “scraper”  -> docs2Txt scraps the words from 3 different types of documents and outputs a list of unique words in each


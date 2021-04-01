# docs2Txt - processes *.doc, *.docs and *.pdf to generate a list of uniqe words
import os
import sys
import time
from pathlib import Path

import docx2txt
import fitz
import win32com.client

# Export *.doc file to text string
def extract_doc_text(doc_file):
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False 
    doc = word.Documents.Open(doc_file)
    text = doc.Range().text    
    doc.Close()
    word.Quit()
    return text

# Export *.docx file to text string
def extract_docx_text(docx_file):
    text = docx2txt.process(docx_file)
    return text

# Export *.pdf file to text string
def extract_pdf_text(pdf_file):
    pdf = fitz.open(pdf_file) 
    text = ""
    for page in pdf:
        text += page.get_text("text")

    pdf.close()
    return text

def create_word_list(text,word_list_file):

    # Split text in to words
    word_list = text.split()
    total_words = len(word_list)

    # Find unique words
    unique_words = set(word_list)
    unique_word_cnt = len(unique_words)

    # Save unique words to text file
    fp = open(word_list_file, "w")
    for word in unique_words:
        try:
            fp.write(str(word) + "\n")
        except: 
            # skip garbage, we do not care about
            continue

    fp.close()
    return(total_words,unique_word_cnt)


def main(argv):

    #Get input file name from command line argument or user imput 
    if len(sys.argv) > 1:
        file_name = argv[0].lower()
    else:
        print ("This program will extract a list of unique words from:")
        print ("MSword documents .doc or docx formats")
        print ("or PDF files.")
        print ("The resulting list of unique words is saved to a file")
        print ("using the same file name, replacing the extension with _word.lst")
        print ('Enter file name including full path:')
        file_name = input() 

    # Make sure it is a valid file
    try:
        size = os.path.getsize(file_name)
    except:
        print ("not a valid file")
        return

    ext = Path(file_name).suffix

    # Call correct library based on extension
    if  ext == ".pdf" :
         text = extract_pdf_text(file_name)
    elif ext == ".docx":
        text = extract_docx_text(file_name)
    elif ext == ".doc":
        text = extract_doc_text(file_name)
    else:
        print("File typenot supported")
    
    # Create output file name = imput file replacing extension with _words.txt
    word_list_file = os.path.splitext(file_name)[0] + "_words.txt"

    # Extract unique list of words from the text and write to text file for procesing by C++ backup program
    total_words, unique_words = create_word_list(text,word_list_file)

    # Print results summary
    print ( file_name + 'Processed')
    print ( 'to ' + word_list_file)
    print ('Total words in Document ' + str(total_words))
    print ('Unique words in Document ' + str(unique_words))
    


if __name__ == "__main__":
   main(sys.argv[1:])


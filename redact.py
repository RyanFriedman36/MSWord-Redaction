
"""

File: redact.py

Project: Automatic legal-document redactionn desktop application

Author: Ryan Friedman

Description: This program will allow a user to supply a Microsoft Word
             Document and a list of proper nouns that need to be redacted.
             It will create a new Word Document with any version of the proper
             nouns replaced with a black highlight. We create a new document to
             ensure fresh metadata.

"""
from docx import Document
from docx.enum.text import WD_COLOR_INDEX
import tkinter as tk
from tkinter import filedialog
from string import punctuation
import os
from os import remove


def binary_search(source, target):
    """ This function searches for a target in source in logarithmic time """

    left = 0
    right = len(source) - 1

    while left <= right:
        median = left + ((right - left) // 2)
        if source[median] == target:
            return True
        elif source[median] < target:
            left = median + 1
        else:
            right = median - 1

    return False


def popupmsg(msg):
    """ This function opens a pop-up window and displays a message. The user
        may click 'okay' to close the window. """

    popup = tk.Tk()
    popup.wm_title("Auto-Redact")
    label = tk.Label(popup, text=msg, font=("Helvetica", 10))
    label.pack(side="top", fill="x", pady=10)
    B1 = tk.Button(popup, text="Okay", command=popup.destroy)
    B1.pack()
    popup.mainloop()


def requestFile():
    """ This function opens a file selection window using tkinter and
    returns the file path """

    root = tk.Tk()
    root.withdraw()

    file_path = filedialog.askopenfilename()

    if not file_path:
        exit(1)

    return file_path


def getDirFromFile(file_path):
    """ This function finds the last forward or back slash in directory path
        and slices the file name from it """

    break_index = -1
    for i in range(len(file_path) - 1, -1, -1):
        if file_path[i] == "\\" or file_path[i] == "/":
            break_index = i
            break

    if break_index != -1:
        file_path = file_path[:break_index + 1]

    return file_path


def processInfoFile(info_path):
    """ Opens the file and cleans/processes information to be redacted
        into the proper state """

    redact_file = open(info_path, "r")
    redact_info = redact_file.read()
    redact_file.close()

    redact_info = redact_info.split(" ")

    for i in range(len(redact_info)):
        redact_info[i] = redact_info[i].strip(",")

    for entry in redact_info:
        if entry == "" or entry == "\n":
            redact_info.remove(entry)
        if " " in entry:
            entry = entry.strip()

    redact_info.sort()

    return redact_info


def processFiles(file_path_1, file_path_2):
    """ file_path_1 is our Word document and file_path_2 is any text file
        with proper nouns listed on separate lines. This is where we create
        our redacted Word File """

    redact_info = processInfoFile(file_path_2)

    doc = Document(file_path_1)

    new_doc = Document()

    for para in doc.paragraphs:
        new_para = processPara(para, redact_info)
        # new_doc.add_paragraph(new_para.text)

    # doc.save(getDirFromFile(file_path_1) + "redacted version.docx")


def processPara(para, redact_info):
    """ Gets indices for words that are to be redacts them then returns a
        a paragraph with those words redacted  """

    redact_indices = getRedactIndices(para, redact_info)
    # print(redact_indices)
    return redact(para, redact_indices)


def getRedactIndices(para, redact_info):
    """ Returns a list containing tuples that represent the indices that any
        instance of redact_info span in para """

    redact_indices = []

    curr_word = ""
    for i in range(len(para.text)):
        char = para.text[i]
        if char == " " or char == "\t" or char == "\n":
            curr_word = curr_word.strip(punctuation)
            if binary_search(redact_info, curr_word):
                redact_indices.append((i - len(curr_word), i - 1))
            curr_word = ""
        else:
            curr_word += char

    # check the final word in each paragraph
    if curr_word:
        curr_word = curr_word.strip(punctuation)
        if binary_search(redact_info, curr_word):
            redact_indices.append((i - len(curr_word) + 1, len(para.text) - 1))

    return redact_indices


def redact(para, redact_indices):
    """ This function modifies our paragraph object with new 'black runs'
        which are our redactions """

    document = Document()
    p = document.add_paragraph()

    for redaction in redact_indices:
        p.add_run()
        document.paragraphs[0].runs[0].font.highlight_color = WD_COLOR_INDEX.BLACK

        start = redaction[0]
        end = redaction[1]
        print(para.text[start:end + 1])

    return "done"

    # temp_para = para

    # document = Document()
    # p = document.add_paragraph()

    # for i in range(len(redact_indices)):
    #     start = redact_indices[i][0]
    #     end = redact_indices[i][1]
    #     my_str = " " * (end - start)
    #     p.add_run(my_str)
    #     document.paragraphs[0].runs[i].font.highlight_color = WD_COLOR_INDEX.BLACK
    #     curr_run = document.paragraphs[0].runs[i]
    #     para.text = para.text[:start]
    #     print(para.text)
    #     para.add_run(curr_run.text)
    #     print(para.text)
    #     para.text = para.text + temp_para.text[end + 2:]
    #     print(para.text)
    # return "done"


def main():


    # popupmsg("Please select the Word document that you would like to redact from.")
    file_path_1 = requestFile()

    file_path_2 = requestFile()

    processFiles(file_path_1, file_path_2)



    # TESTING:

    # my_doc = Document()
    # for i in range(10):
    #     my_doc.add_paragraph(str(i))


    # my_doc.paragraphs[1] = my_doc.paragraphs[1].add_run("adhfalsdfj").bold = True
    # my_doc.paragraphs[1].runs[-1].font.highlight_color = WD_COLOR_INDEX.BLACK
    # my_doc.paragraphs[1].runs[-1].bold = False

    # for para in my_doc.paragraphs:
    #     print(para.text)

    # target_dir = getDirFromFile(requestFile())
    # os.remove(target_dir + "output_test.docx")

    # my_doc.save(target_dir + "output_test.docx")




    # popupmsg("The redacted version of your file has been created.")

    # display directory?

    # ask if they would like to open it in word

    # ask if they would like you to port to a pdf?


if __name__ == "__main__":
    main()

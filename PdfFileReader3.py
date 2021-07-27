
"""
FileName: PdfFileReader3.py
Author: Clint Kline
Created: 7-18-2021
Last Updated: 7-27-2021
Purpose:
    - copy text from a PDF file into a word docx

    - An offset variable is provided to compensate for the fact that
    the PDF reader includes all pages before the first numbered page
    in its total page count.
    (this includes contents, forward, preface, everything except the
    cover, ignore the cover.)

    - To determine the offset variable, count all pages before the
    first numbered page except the cover, and enter that number
    when prompted.

>>> Issues: 
    - If the user chooses to back up after a page number is given to copy a single page,
    the 'reader' object is converted from an object, to a string. The cause of this issue
    is unidentified as of 07-27-2021. 


"""


import PyPDF2
from docx import Document
from tkinter.filedialog import askdirectory
offset = 0


#####################################
# MAIN
#####################################


def main(reader, offset):
    oneOrMany(reader, offset)


######################################
# BACK FUNCTION
######################################


def back(var1, func, p1, *args):
    if var1.lower() == "b":
        func(p1, *args)


######################################
# TESTFILENAME FUNCTION
######################################


def testFileName(var1, func, output):
    # provide a list of illegal filname characters
    illegalChars = ["`", "~", "!", "@", "#", "$", "%", "^", "&", "*", "+", "=", "'", "|", ":", ";",
                    "[", "]", "{", "}", "<", ">", ",", "/" "?", "\\", "\"", "\n", "\r", "\t", "\b", "\f"]
    # test if name contains any of the characters in illegalChars and assign results to legal
    FileNameIsLegal = [char for char in illegalChars if (char in var1)]
    # vvv print out the value of var legal
    # print(str(bool(legal)))

    if str(bool(FileNameIsLegal)) == "True":
        print("\nSpecial characters cannot be used to name a file.")
        back("b", func, output)
        return 1
    elif var1.startswith(tuple('0123456789')):
        print("\nFile names cannot start with a number.")
        back("b", func, output)
        return 2
    elif var1.startswith(tuple('._-')):
        print("Start a filename with letters only please.")
        back("b", func, output)
        return 3
    elif len(var1) > 31:
        print("\nPlease keep file names under 31 characters.")
        back("b", func, output)
        return 4
    else:
        return 5


######################################
# ONE OR MANY FUNCTION
######################################


def oneOrMany(reader, offset):
    # user chooses the pages of the pdf to copy
    oom = str(input(
        "\nDo you want to copy more than one page? y/n.\n(you can hit \"b\" at any time to go back.)\n"))

    back(oom, offsetFunc, reader)

    if oom == "Y" or oom == "y":
        firstPageFunc(reader, offset)
    elif oom == "N" or oom == "n":
        one(reader, offset)
    else:
        print("\n\nPlease enter \"y\", \"n\", or \"b\"")
        oneOrMany(reader, offset)


######################################
# ONE PAGE FUNCTION
######################################


def one(reader, offset):
    pageNum = input("\nWhich page do you want to copy? ")

    if pageNum.isdigit() == True:
        intPageNum = int(pageNum)
        # test reader value
        print(reader)
        # set pageObj value to page number passed to reader variable
        pageObj = reader.getPage(intPageNum + offset)
        output = pageObj.extractText()
        format(output, reader, offset)
    elif pageNum.isdigit() == False:
        back(pageNum, oneOrMany, reader, offset)
    else:
        print(
            "Input must be either an integer or the letter \"b\" to go back.")
        one(reader, offset)


#######################################
# firstPageFunc FUNCTION
#######################################

def firstPageFunc(reader, offset):
    firstPage = input("\nFirst Page:\n")
    # hit b to go back to oneOrMany
    if firstPage.isdigit() == False:
        back(firstPage, oneOrMany, reader, offset)
    # test to ensure input is int type
    elif firstPage.isdigit() == True:
        intFirstPage = int(firstPage)
        # accept number of last page to be copied
        lastPageFunc(reader, offset, intFirstPage)
    # if firstPage input is not an int
    else:
        print("please enter a page number.")
        firstPageFunc(reader, offset)

#######################################
# lastPageFunc FUNCTION
#######################################


def lastPageFunc(reader, offset, intFirstPage):
    lastPage = input("\nLast Page:\n")  # 20
    # hit b to restart firstPageFunc
    if lastPage.isdigit() == False:
        back(lastPage, firstPageFunc, reader, offset)
    # test to ensure input is int type
    elif lastPage.isdigit() == True:
        intLastPage = int(lastPage)
        # ensure the value of firstPage is < the value of lastPage
        if intFirstPage < intLastPage:
            # set variables
            output = ""
            pageArray = []
            # create an array of all pages to be copied
            while intFirstPage <= intLastPage:
                pageNum = intFirstPage
                pageArray.append(pageNum)
                intFirstPage += 1
            print("\nPages to be copied: ", pageArray)
            # extract text from multiple pages
            for page in pageArray:
                pageObj = reader.getPage(page + offset)
                output += pageObj.extractText()
            # send output to format() function
            format(output, reader, offset)
        # if the value of lastPage is > firstPage
        else:
            print("\n\nThe first page must come before the last page.")
            firstPageFunc(reader, offset)
    # if lastPage input is not an int
    else:
        print("please enter a page number.")
        firstPageFunc(reader, offset)


####################################
# OFFSET DETERMINATION
####################################


def offsetFunc(reader, *args):
    # accept an offset
    """
    offset = int(input("\nIf there is an offset of the page number enter it here, if not enter 0:\n(- To determine the offset variable, count all pages before the first numbered page (this includes contents, forward, preface, everything except the cover.))\n"))
    main(reader, offset)

    """
    try:
        offset = int(input("\nIf there is an offset of the page number enter it here, if not enter 0:\n(- To determine the offset variable, count all pages before the first numbered page (this includes contents, forward, preface, everything except the cover.))\n"))
        main(reader, offset)
    except Exception as e:
        print("\nThat didnt work. Input must be of type integer.\nError: ", e, "\n\n")
        offsetFunc(reader)

######################################
# FORMATTING
######################################


# corrects odd character conversions and indentations

def format(output, reader, offset):
    # account for S characters that rep " - " symbols
    noDashOutput = output.replace("Š", " - ")
    # remove unwanted auto-line breaks
    noDashOutputRemovebadLineBreaks = noDashOutput.replace("\n", "")
    # enter line breaks
    noDashOutputWithGoodLineBreaks = noDashOutputRemovebadLineBreaks.replace(
        ".", ".\n\n")
    tradeMarkToApostrophe = noDashOutputWithGoodLineBreaks.replace("™", "'")
    saveAsDocx(tradeMarkToApostrophe, reader, offset)


######################################
# OUTPUT & FILE NAMING
######################################


def saveAsDocx(output, *args):

    #######
    # vvv see the output as its written
    # print(noDashOutputWithGoodLineBreaks)
    #######

    ######################################
    # SAVE OUTPUT TO A WORD DOC (DOCX)
    ######################################
    # create a new word document
    docx = Document()
    # add the selected contents to the document
    docx.add_paragraph(output)

    # ask user for a file name to save as .docx
    name = input("\nWhat do you want to name the docx file?\n")
    # remove any whitespace from the name. deal with it.
    name = name.strip()

    back(name, one, output, offset)

    if testFileName(name, saveAsDocx, output) == 1 or 2 or 3:
        # open file dialog to select save location
        saveLocation = askdirectory(parent=None)
        # concatenate selected file path to designated file name
        fileToLocation = saveLocation + "/" + name + ".docx"
        # save file @ location
        docx.save(fileToLocation)
        # print save confirmation
        print("\n\nDocument saved as", name + ".docx @", saveLocation)
    elif testFileName(name, saveAsDocx, output) == 4 or 5:
        print(name, "was not saved.")
        one(reader, *args)


###############################
# INITIALIZATION
###############################
if __name__ == "__main__":
    # open(create) a pdf object, must enter complete filepath
    print("\n\n" + "*" * 50)
    print("*** Clint's PDF to DOCX ***")
    print("*" * 50)

    fileName = input(
        "\n\nPlease enter the complete filepath to the PDF you wish to copy from: ")
    try:
        pdfObj = open(fileName, 'rb')
    except Exception as e:
        print("\nThat didnt work. Error: ", e, "\n\n")

    # print fileName to confirm correct files
    print("\nObjectID: ", pdfObj)
    # create a pdf reader object
    reader = PyPDF2.PdfFileReader(pdfObj)
    # print the total number of pages
    print("\nNumber of Pages: ", reader.numPages, "\n")
    # send reader object to the offset function
    offsetFunc(reader)

    #################################
    # CLOSE PDF FILE
    #################################

    pdfObj.close()
    print("\n\nClosed File:", fileName, "\n")

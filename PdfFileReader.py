"""
FileName: PdfFileReader.py
Author: Clint Kline
Created: 7-18-2021
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

"""

import PyPDF2
from docx import Document


#####################################
# MAIN
#####################################

def main(reader, offset):
    oneOrMany(reader, offset)

######################################
# ONE OR MANY FUNCTION
######################################


def oneOrMany(reader, offset):
    # user chooses the pages of the pdf to copy
    oom = str(input(
        "\nDo you want to copy more than one page? y/n.\n(you can hit \"b\" at any time to go back.)\n"))
    if oom == "Y" or oom == "y":
        many(reader, offset)
    elif oom == "N" or oom == "n":
        one(reader, offset)
    elif oom == "B" or oom == "b":
        offsetFunc(reader)
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
        pageObj = reader.getPage(intPageNum + offset)
        output = pageObj.extractText()
        format(output)
    elif pageNum == "B" or pageNum == "b":
        oneOrMany(reader, offset)
    else:
        print(
            "Input must be either an integer or the letter \"b\" to go back.")
        one(reader, offset)


#######################################
# MANY PAGES FUNCTION
#######################################

def many(reader, offset):
    # accept number of first page to be copied
    firstPage = input("\nFirst Page:\n")
    # hit b to go back to oneOrMany
    if firstPage == "B" or firstPage == "b":
        oneOrMany(reader, offset)
    # test to ensure input is int type
    elif firstPage.isdigit() == True:
        intFirstPage = int(firstPage)
        # accept number of last page to be copied
        lastPage = input("\nLast Page:\n")  # 20
        # hit b to restart many
        if lastPage == "B" or lastPage == "b":
            many(reader, offset)
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
                format(output)
            # if the value of lastPage is > firstPage
            else:
                print("\n\nThe first page must come before the last page.")
                many(reader, offset)
        # if lastPage input is not an int
        else:
            print("please enter a page number.")
            many(reader, offset)
    # if firstPage input is not an int
    else:
        print("please enter a page number.")
        many(reader, offset)

####################################
# OFFSET DETERMINATION
####################################


def offsetFunc(reader):
    # accept an offset
    try:
        offset = int(input("\nIf there is an offset of the page number enter it here, if not enter 0:\n(- To determine the offset variable, count all pages before the first numbered page (this includes contents, forward, preface, everything except the cover.))\n"))
        main(reader, offset)
    except Exception as e:
        print("\nThat didnt work. Input must be of type integer.\nError: ", e, "\n\n")


######################################
# FORMATTING
######################################

# corrects odd character conversions and indentations

def format(output):
    # account for S characters that rep " - " symbols
    noDashOutput = output.replace("Š", " - ")
    # remove unwanted auto-line breaks
    noDashOutputRemovebadLineBreaks = noDashOutput.replace("\n", "")
    # enter line breaks
    noDashOutputWithGoodLineBreaks = noDashOutputRemovebadLineBreaks.replace(
        ".", ".\n\n")
    tradeMarkToApostrophe = noDashOutputWithGoodLineBreaks.replace("™", "'")
    saveAsDocx(tradeMarkToApostrophe)


######################################
# OUTPUT
######################################

def saveAsDocx(output):
    # get text from page

    # remove '#' from next line if you wish to see the output as its written
    # print(noDashOutputWithGoodLineBreaks)

    ######################################
    # SAVE OUTPUT TO A WORD DOC (DOCX)
    ######################################
    # create a new word document
    docx = Document()
    # add the selected contents to the document
    docx.add_paragraph(output)
    # save docx file
    # name file
    name = input("\nWhat do you want to name the docx file?\n")
    if name == "B" or name == "b":
        offsetFunc(reader)
    # save file
    docx.save(name + ".docx")
    print("\n\nDocument saved as", name + ".docx")


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

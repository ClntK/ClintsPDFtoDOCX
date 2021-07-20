"""
FileName: PdfFileReader.py
Author: Clint Kline
Created: 7-18-2021
Purpose:
    - copy text from a PDF file into a word docx

    - An offset variable is provided to compensate for the fact that
    the PDF reader includes all pages before the first numbered page 
    (this includes contents, forward, preface, everything except the 
    cover, ignore the cover.)

    - To determine the offset variable, count all pages before the 
    first numbered page except the cover, and enter that number 
    when prompted.

"""

import PyPDF2
from docx import Document


def main(reader, offset):
    # user chooses the pages of the pdf to copy
    oneOrMany = str(input(
        "\nDo you want to copy more than one page? y/n.\n(you can hit \"b\" at any time to re-enter data.)\n"))
    if oneOrMany == "Y" or oneOrMany == "y":
        many(reader, offset)
    elif oneOrMany == "N" or oneOrMany == "n":
        one(reader, offset)
    elif oneOrMany == "B" or oneOrMany == "b":
        oneOrMany(reader, offset)
    else:
        print("please enter \"y\", \"n\", or \"b\"")

# extract text from one page


def one(reader, offset):
    pageNum = int(input("Which page? #"))
    if pageNum == "B" or pageNum == "b":
        one()
    elif isinstance(pageNum, int) == True:
        pageObj = reader.getPage(pageNum + offset)
        output = pageObj.extractText()
        format(output)
    else:
        print("please enter a number.")

# extract text from many pages.


def many(reader, offset):
    # accept number of first page to be copied
    firstPage = int(input("\nFirst Page:\n"))  # 10
    # hit b to restart many
    if firstPage == "B" or firstPage == "b":
        many()
    # test to ensure input is int type
    elif isinstance(firstPage, int) == True:
        # accept number of last page to be copied
        lastPage = int(input("\nLast Page:\n"))  # 20
        # hit b to restart many
        if lastPage == "B" or lastPage == "b":
            many()
        # test to ensure input is int type
        elif isinstance(lastPage, int) == True:
            # set variables
            output = ""
            pageArray = []
            # create an array of all pages to be copied
            while firstPage <= lastPage:
                pageNum = firstPage
                pageArray.append(pageNum)
                firstPage += 1
            print("\nPages to be copied: ", pageArray)
            # extract text from multiple pages
            for page in pageArray:
                pageObj = reader.getPage(page + offset)
                output += pageObj.extractText()
            # send output to format() function
            format(output)
        # if input is not an int
        else:
            print("please enter a number.")
    else:
        print("please enter a number.")


######################################
# FORMATTING
######################################

def format(output):
    # account for S characters that rep " - " symbols
    noDashOutput = output.replace("Š", " - ")
    # remove unwanted auto-line breaks
    noDashOutputRemovebadLineBreaks = noDashOutput.replace("\n", "")
    # enter line breaks
    noDashOutputWithGoodLineBreaks = noDashOutputRemovebadLineBreaks.replace(
        ".", ".\n\n")
    saveAsDocx(noDashOutputWithGoodLineBreaks)


######################################
# OUTPUT
######################################
def saveAsDocx(noDashOutputWithGoodLineBreaks):
    # get text from page

    # remove '#' from next line if you wish to see the output as its written
    # print(noDashOutputWithGoodLineBreaks)

    ######################################
    # SAVE OUTPUT TO A WORD DOC (DOCX)
    ######################################
    # create a new word document
    docx = Document()
    # add the selected contents to the document
    docx.add_paragraph(noDashOutputWithGoodLineBreaks)
    # save docx file
    # name file
    name = input("\nWhat do you want to name the docx file?\n")
    # save file
    docx.save(name + ".docx")
    print("\n\nDocument saved as", name + ".docx")


if __name__ == "__main__":
    # open(create) a pdf object, must enter complete filepath
    print("\n\n" + "*" * 50)
    print("*** Clint's PDF to DOCX ***")
    print("*" * 50)

    fileName = input(
        "\n\nPlease enter the complete filepath to the PDF you wish to copy from: ")
    pdfObj = open(fileName, 'rb')
    # print fileName to confirm correct files
    print("\nObjectID: ", pdfObj)
    # create a pdf reader object
    reader = PyPDF2.PdfFileReader(pdfObj)
    # print the total number of pages
    print("\nNumber of Pages: ", reader.numPages, "\n")

    # accept an offset
    offset = int(input("If there is an offset of the page number enter it here:\n(- To determine the offset variable, count all pages before the first numbered page (this includes contents, forward, preface, everything except the cover.))\n"))

    if isinstance(offset, int) == True:
        # begin process
        main(reader, offset)
    else:
        print("please enter a number:\n")

    # close file
    pdfObj.close()
    print("\n\nClosed File:", fileName, "\n")

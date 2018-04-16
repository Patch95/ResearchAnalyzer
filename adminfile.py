#!/usr/bin/python3
import shutil
import os


#  Function that reads the complete file
def readfile(path):
    # First of all we must open the file that we are going to read.
    # Use 'rb' instead of 'r' if it is a binary file.
    infile = open(path, 'r')

    file = infile.read()
    # print(infile.read())
    # Close the file.
    infile.close()
    return file


# Function writes to a file (overwrite)
def writefile(path, datos):
    outfile = open(path, 'w')  # We indicate the value 'w'.
    outfile.write(datos)
    outfile.close()


# Function to create folders
def createfolder(path):
    if(False == (os.path.exists(path))):  # if the address does not exist
        os.mkdir(path)


# Function that copies a file
def copyfile(pathorigin, pathdestiny):
    shutil.copy(pathorigin, pathdestiny)


def movefile(pathorigin, pathdestiny):
    shutil.move(pathorigin, pathdestiny)


def deletefile(path):
    os.remove(path)



# Pdf-picture-convert-to-word
To convert the PDF documents to pictures, then insert these pictures into an MS-word document in sequence


--------
# Features
--------
- In a Windows environment, batch convert PDF files to images and store them in individual folders.
- To insert the images from each folder into a Word document in sequence.

# Requirements
------------
-windows10 x64, windows7 x64 
-Python 3.10

# Installation
------------
## A.Install and deploy environment python 3.10


## B.Config
1. make Path --> C:\Program Files\OneClickPDF
2. copy PdftoWord.py into oneClickPDF folder
3. create a shortcut for 
4. move the shortcut of PdftoWord.py to this path C:\Users\***\AppData\Roaming\Microsoft\Windows\SendTo
5. modify function Menus of shortcut to this:
     - Target:   C:\Users\***\AppData\Local\Programs\Python\Python310\python.exe "C:\Program Files\OneClickPDF\PdftoWord.py "
     - Star in:  "C:\Program Files\OneClickPDF"
     - *Note that the Python installation path is the default installation path
 
6. save config

## C. Install necessary library files or packages

run :

    pip3 install -r requirements.txt  


# How to run it
-------------
Select multiple or one document with the left mouse, right-click, and click the “PdftoWord” button in the “send to” menu in the context menus, and then， the program starts to run.





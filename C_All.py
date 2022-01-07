#!C:\Users\kenno\AppData\Local\Programs\Python\Python38\python.exe
from win32com import client
import os, glob
import itertools as it
import asyncio
import cgi
def ImportCss():
    print("Content-Type: text/html\n")
    print(f'<link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css"> <script src="https://ajax.googleapis.com/ajax/libs/jquery/3.5.1/jquery.min.js"></script><script src="https://cdnjs.cloudflare.com/ajax/libs/popper.js/1.16.0/umd/popper.min.js"></script><script src="https://maxcdn.bootstrapcdn.com/bootstrap/4.5.2/js/bootstrap.min.js"></script>')
def ButtonPrinter(self, scriptPath):
    print(f'<h3 class="title pt-xl-5"><a href={scriptPath}>{self}</a></h3>')
# Funkcja odpowiadająca za konwersję wszystkich plików excelowskich znajdujących się w folderze Conventer.
def ConvertAll():
    localPath = os.getcwd()
    app = client.DispatchEx("Excel.Application")
    app.Interactive = False
    app.Visible = False
    path = localPath + r'/Conventer/'
    os.chdir(path)
    def multiple_file_types(*patterns):
        return it.chain.from_iterable(glob.iglob(pattern) for pattern in patterns)
    if not multiple_file_types("*.xlsx", "*.xlsm", "*.xls"):
        ImportCss()
        print("W folderze nie ma plików excela przeznaczonych do konwersji. Proszę przenieść do folderu pliki przeznaczone do konwersji.")
    else:
        for file in multiple_file_types("*.xlsx", "*.xlsm", "*.xls"):
            inputFile = os.path.join(path + file)
            if file in glob.glob("*.xls"):
                length = len(inputFile)
                length -= 4
                outputFile = inputFile[:length]
                Workbook = app.Workbooks.Open(inputFile)
            else:    
                length = len(inputFile)
                length -= 5
                outputFile = inputFile[:length]
                Workbook = app.Workbooks.Open(inputFile)
            try:
                Workbook.ActiveSheet.ExportAsFixedFormat(0, outputFile)
            except Exception as e:
                print("Konwersja na PDF nie udala sie.")
                print(str(e))
            finally:
                Workbook.Close()
                app.Quit()
        ImportCss()
        print('<div class="jumbotron pt-5 text-center"><h1 class="title">Konwersja plików zostala zakończona</h1></div>')
        print('<div class="container text-center">')
        print('Kliknij w napis poniżej, aby wyświetlić listę plików')
        ButtonPrinter("Lista plików","PDFlist.py")
        print("</div>")
ConvertAll()
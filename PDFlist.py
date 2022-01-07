#!C:\Users\kenno\AppData\Local\Programs\Python\Python38\python.exe
# Wymaga biblioteki pywin32 i zainstalowanego Excela
# Aby skrypt działał w HTML w pierwszej linijce kodu potrzebna jest ścieżka do pythona
from sys import path
import os, glob
def ShowlistInHtml():
    localPath = os.getcwd()
    mainPath = localPath + r'/Conventer/'
    os.chdir(mainPath)
    if not glob.glob("*.pdf"):
        print("W folderze nie ma żadnego pliku PDF")
    else:
        print('<div class="jumbotron pt-5 text-center">')
        print('<h1 class="title">Lista plików pdf: </h1>')
        print('</div>')
        print('<div class="container">')
        print('<p class="title text-center"><i>Kliknij na nazwę, aby wyświetlić lub pobrać plik</i></p>')
        print('<ul class="list-group">')
        for pdfFiles in glob.glob("*.pdf"):
            downloadFile = '/Conventer/' + pdfFiles
            print(f'<li class="list-group-item"><a href={downloadFile}>{pdfFiles}</a></li>', end='\n')
        print('</ul>')
        print("</div>")
def ImportCss():
    print("Content-Type: text/html\n")
    print(f'<link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css"> <script src="https://ajax.googleapis.com/ajax/libs/jquery/3.5.1/jquery.min.js"></script><script src="https://cdnjs.cloudflare.com/ajax/libs/popper.js/1.16.0/umd/popper.min.js"></script><script src="https://maxcdn.bootstrapcdn.com/bootstrap/4.5.2/js/bootstrap.min.js"></script>')
ImportCss()
ShowlistInHtml()

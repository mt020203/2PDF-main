#!C:\Users\kenno\AppData\Local\Programs\Python\Python38\python.exe
# Aby skrypt działał w HTML w pierwszej linijce kodu potrzebna jest ścieżka do pythona
import os, glob
import itertools as it
def ImportCss():
    print("Content-Type: text/html\n")
    print(f'<link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css"> <script src="https://ajax.googleapis.com/ajax/libs/jquery/3.5.1/jquery.min.js"></script><script src="https://cdnjs.cloudflare.com/ajax/libs/popper.js/1.16.0/umd/popper.min.js"></script><script src="https://maxcdn.bootstrapcdn.com/bootstrap/4.5.2/js/bootstrap.min.js"></script>')
def GenerateLinks(ipAddress):
        localPath = os.getcwd()
        path = localPath + r'/Conventer/'
        os.chdir(path)
        def multiple_file_types(*patterns):
            return it.chain.from_iterable(glob.iglob(pattern) for pattern in patterns)
        open('links.txt', "w")
        for file in multiple_file_types("*.xlsx", "*.xlsm", "*.xls"):
            textFile = open('links.txt', "a")
            textFile.write(f'{ipAddress}/?filename={file}\n')
        ImportCss()
        print("Linki zostały wygenerowane")
GenerateLinks("localhost")
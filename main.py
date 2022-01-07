#!C:\Users\kenno\AppData\Local\Programs\Python\Python38\python.exe
# Wymaga biblioteki pywin32 i zainstalowanego Excela
# Aby skrypt działał w HTML w pierwszej linijce kodu potrzebna jest ścieżka do pythona
from win32com import client
import os, glob
import itertools as it
import asyncio
import cgi
# Przekierowanie zmiennych z JS do Pythona, przesłanych przez POST
form = cgi.FieldStorage()
fileName =  form.getvalue('filename')
fileName = str(fileName)
def ImportCss():
    print("Content-Type: text/html\n")
    print(f'<link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css"> <script src="https://ajax.googleapis.com/ajax/libs/jquery/3.5.1/jquery.min.js"></script><script src="https://cdnjs.cloudflare.com/ajax/libs/popper.js/1.16.0/umd/popper.min.js"></script><script src="https://maxcdn.bootstrapcdn.com/bootstrap/4.5.2/js/bootstrap.min.js"></script>')
# Funkcja odpowiadająca za konwersję danego pliku z parametru w linku.
def ConvertSpecific(self):
    localPath = os.getcwd()
    app = client.DispatchEx("Excel.Application")
    app.Interactive = False
    app.Visible = False
    path = localPath + "\\Conventer\\"
    os.chdir(path)
    def multiple_file_types(*patterns):
        return it.chain.from_iterable(glob.iglob(pattern) for pattern in patterns)
    file = str(self)
    # Walidacja plików
    if not multiple_file_types("*.xlsx", "*.xlsm", "*.xls"):
        ImportCss()
        print('<div class="d-flex flex-wrap align-content-center justify-content-center text-danger container text-center align-items-center"><h1>W folderze nie ma plików excela przeznaczonych do konwersji. Proszę przenieść do folderu pliki przeznaczone do konwersji.</h1></div>')
    else:
    # Usunięcie zbędnych rozszerzeń plików
        inputFile = os.path.join(path + file)
        if file in glob.glob("*.xls"):
            length = len(inputFile)
            length -= 4
            outputFile = inputFile[:length]
            Workbook = app.Workbooks.Open(inputFile)
            # Usunięcie zbędnych rozszerzeń plików dla wyświetlania na stronie
            if self.endswith('.xlsx') or self.endswith('.xlsm'): 
                viewFile = len(self)
                viewFile -= 5
                viewFile = self[:viewFile]
            elif self.endswith('.xls'):
                viewFile = len(self)
                viewFile -= 4
                viewFile = self[:viewFile]
                viewFile = str(viewFile)
        elif file in glob.glob("*.xlsx") or file in glob.glob("*.xlsm"):    
            length = len(inputFile)
            length -= 5
            outputFile = inputFile[:length]
            Workbook = app.Workbooks.Open(inputFile)
            # Usunięcie zbędnych rozszerzeń plików dla wyświetlania na stronie
            if self.endswith('.xlsx') or self.endswith('.xlsm'): 
                viewFile = len(self)
                viewFile -= 5
                viewFile = self[:viewFile]
            elif self.endswith('.xls'):
                viewFile = len(self)
                viewFile -= 4
                viewFile = self[:viewFile]
                viewFile = str(viewFile)
        else:
            ImportCss()
            print('<div class="d-flex flex-wrap align-content-center justify-content-center text-danger container text-center align-items-center bg-dark"><h1>Podany link jest niepoprawny. Sprawdź poprawność linku.</h1></div>')
        try:
            Workbook.ActiveSheet.ExportAsFixedFormat(0, outputFile)
        except Exception as e:
            print('<div class="d-flex flex-wrap align-content-center justify-content-center text-danger container text-center align-items-center bg-dark"><h1>Konwersja do PDF nie udała się.</h1></div>')
        finally:
            Workbook.Close()
            app.Quit()
            # Wyświetlanie pliku na stronie
            print("Content-Type: text/html\n")
            print("<head>")
            print('<script type="text/javascript" src="https://code.jquery.com/jquery-1.4.3.min.js"></script>')
            print('''
                    <script type="text/javascript">
                    $(window).load(function(){
                    $('#pdf').attr('src','''f'{viewFile}'''');
                    
                    });
                    </script> ''')
            print('</head>')
            print('<body>')
            print(f'<iframe src="\Conventer\{viewFile}.pdf" height="100%" width="100%" id="pdf"></iframe>')
            print('</body>')
ConvertSpecific(fileName)

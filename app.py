from flask import Flask, render_template, request, redirect, url_for
import os
import win32api
import win32print
import tempfile

app = Flask(__name__)

# Dossier pour les fichiers téléchargés
UPLOAD_FOLDER = 'uploads'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# Assurez-vous que le dossier existe
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return redirect(request.url)
    
    file = request.files['file']
    if file.filename == '':
        return redirect(request.url)
    
    if file:
        # Sauvegarde le fichier
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
        file.save(filepath)
        
        # Lit le nombre de copies
        try:
            copies = int(request.form['copies'])
        except ValueError:
            copies = 1

        # Imprime le fichier
        print_file(filepath, copies=copies)
        
        return redirect(url_for('index'))

def print_file(file_path, copies=1, printer_name=None):
    if printer_name is None:
        printer_name = win32print.GetDefaultPrinter()
    
    if not os.path.exists(file_path):
        print(f"Le fichier {file_path} n'existe pas.")
        return
    
    try:
        hPrinter = win32print.OpenPrinter(printer_name)
        try:
            devmode = win32print.GetPrinter(hPrinter, 2)["pDevMode"]
            devmode.Copies = copies
            pDevMode = devmode
            pinfo = win32print.GetPrinter(hPrinter, 2)
            pinfo["pDevMode"] = pDevMode
            pinfo["pSecurityDescriptor"] = None
            hJob = win32print.StartDocPrinter(hPrinter, 1, (
                "Impression Automatique",
                None,
                "RAW"
            ))

            try:
                win32print.StartPagePrinter(hPrinter)
                win32api.ShellExecute(
                    0,
                    "printto",
                    file_path,
                    f'"{printer_name}"',
                    ".",
                    0
                )
            finally:
                win32print.EndDocPrinter(hPrinter)

        finally:
            win32print.ClosePrinter(hPrinter)

    except Exception as e:
        print(f"Une erreur est survenue lors de l'impression : {e}")

if __name__ == '__main__':
    app.run(debug=True)

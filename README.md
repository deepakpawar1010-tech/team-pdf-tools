# Team PDF Tools — Python Edition

This is a local Flask version of the PDF tool, with a separate HTML page and a `project.json` configuration file.

## Run locally (Command Prompt / CMD)

```cmd
cd /d "c:\Users\Deepak Pawar\Downloads\Team_PDF_Tools_Python_range"
python -m venv .venv
.venv\Scripts\activate.bat
pip install -r requirements.txt
python app.py
```


Open `http://127.0.0.1:5050`.

## Privacy and deployment

The server does not write uploads or results to disk. When deployed to a shared server, team PDFs travel to that server for processing, so use a trusted private host (for example an internal company server) if documents are confidential.

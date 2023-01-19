FROM python:3.10

ADD sheets.py /app/sheets.py
ADD functions.py /app/functions.py
ADD creds.json /app/creds.json
ADD months_cz.py /app/months_cz.py
ADD invoices /app/invoices
ADD reports /app/reports

WORKDIR /app

RUN pip install gspread google-api-python-client python-dateutil oauth2client requests

ENTRYPOINT [ "python", "./sheets.py" ]

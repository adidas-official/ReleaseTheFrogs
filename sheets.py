from oauth2client.service_account import ServiceAccountCredentials
from googleapiclient.discovery import build
from pathlib import Path
import gspread
import logging
import functions
import time

logging.basicConfig(
    filename='run.log',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S')

SCOPE = ["https://spreadsheets.google.com/feeds",
         "https://www.googleapis.com/auth/spreadsheets",
         "https://www.googleapis.com/auth/drive.file",
         "https://www.googleapis.com/auth/drive"]

CREDS = ServiceAccountCredentials.from_json_keyfile_name('creds.json', SCOPE)
ACCESS_TOKEN = CREDS.get_access_token().access_token
CLIENT = gspread.authorize(CREDS)
DRIVE_SERVICE = build('drive', 'v3', credentials=CREDS)


def main():

    sheets = functions.get_all_sheets(DRIVE_SERVICE)
    latest_report = functions.get_latest_report(sheets)
    if latest_report[0]:
        logging.info('Checking for number of hours [REMOTE]')
        hours = functions.get_hours(CLIENT.open(latest_report[0]))
        if hours != '0':

            logging.info('Checking for existing reports [REMOTE]')
            new_report_name = functions.new_report_name()
            # logging.info(new_report_name)
            if new_report_name not in list(sheets.keys()):
                functions.make_new_report(CLIENT, new_report_name)
                functions.rename_form(DRIVE_SERVICE, new_report_name)
            else:
                logging.info(f'Report {latest_report[0]} already exists [REMOTE]')

            functions.export_last_report(latest_report, DRIVE_SERVICE)
            latest_invoice = functions.get_latest_invoice(sheets.keys())

            functions.prepare_invoice(CLIENT, latest_report[0], latest_invoice)
            functions.export_latest_invoice(CLIENT, ACCESS_TOKEN, DRIVE_SERVICE)

        else:
            logging.warning('No hours in last report [REMOTE]')

    else:
        functions.make_new_report(CLIENT, functions.new_report_name())
        time.sleep(5)
        functions.rename_form(DRIVE_SERVICE)

    if Path('run.log').exists():
        with open('run.log','a') as log:
            log.write('\n')


if __name__ == '__main__':
    main()

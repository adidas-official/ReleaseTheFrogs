from oauth2client.service_account import ServiceAccountCredentials
from googleapiclient.discovery import build
import gspread
import re
import requests
import logging
import months_cz
from datetime import datetime, timedelta
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s', datefmt='%Y-%m-%d %H:%M:%S')


scope = ["https://spreadsheets.google.com/feeds",
         "https://www.googleapis.com/auth/spreadsheets",
         "https://www.googleapis.com/auth/drive.file",
         "https://www.googleapis.com/auth/drive"]

creds = ServiceAccountCredentials.from_json_keyfile_name('creds.json', scope)
client = gspread.authorize(creds)
# sheet = client.open('Copy of template-report-naturaservis')

drive_service = build('drive', 'v3', credentials=creds)


def get_all_sheets(service):
    all_sheets = service.files().list(
        q='mimeType=\'application/vnd.google-apps.spreadsheet\'',
        fields='files(name)').execute()[
        'files']
    sheet_names = []
    for s in all_sheets:
        sheet_names.append(s['name'])
    return sheet_names


def get_latest_report(files):
    latest, latest_name = 0, 0

    for f in files:
        if 'pracovni vykaz' in str(f):
            number = int(f[:2])
            if number > latest:
                latest = number
                latest_name = f

    return latest_name


def start_of_month(dt):
    todays_month = dt.month
    yesterdays_month = (dt + timedelta(days=-1)).month
    return True if yesterdays_month != todays_month else False


# get name of current month
# use it for creating name of new spreadsheet
# returns number[string with leading zero] of month, and it's name
def get_month_name():
    month_now = datetime.now().month
    m_num = month_now % 12
    m_name = months_cz.months_cz[m_num - 1]
    m_num = str(m_num).zfill(2)
    return m_num, m_name


# get new name for template
# returns something like 01-pracovni vykaz leden, depending on current date
def new_report_name():
    month_num, month_name = get_month_name()
    new_spreadsheet_name = f'{month_num}-pracovni vykaz {month_name}'
    return new_spreadsheet_name


# fill b4 cell with correct month name
def make_new_report(authorized_client, report_name):
    template_report_id = '1Cpexu6iBaKRvDCO7uBV6dQr0-t3qfQQgPUDipGJvBcE'
    report = authorized_client.copy(template_report_id, title=report_name, copy_permissions=True)
    year = datetime.now().year
    report.sheet1.update('b4', f'měsíc: {get_month_name()[1].capitalize()} {year}')


def get_hours(report):
    return report.sheet1.acell('G6').value


def get_latest_invoice(all_sheets):

    last = 0
    sheet_name_regex = re.compile(r'\d{8}')
    for sheet_name in all_sheets:
        if sheet_name_regex.search(sheet_name):
            number = int(sheet_name[4:])
            if number > last:
                last = number
    return last


# returns something like 20220002
def new_invoice_name(latest_invoice):
    year = datetime.now().year
    new_spreadsheet_name = str(year) + str(int(latest_invoice + 1)).zfill(4)
    return new_spreadsheet_name


new_report = new_report_name()
reports = get_all_sheets(drive_service)

latest_invoice_num = get_latest_invoice(get_all_sheets(drive_service))
invoice_name = new_invoice_name(latest_invoice_num)


def prepare_invoice(invoice_num):
    new_invoice = client.copy('1HJIUHBL2kRhS69OTKAfQSYRVSx0_3W1SwFLIdHaUylA', invoice_num, copy_permissions=True)
    hours = get_hours(client.open(get_latest_report(reports)))

    # formating total cost to format 12 345,00
    total = 150 * float(hours)
    total = f'{total:,.2f}'
    total = total.replace(',', ' ').replace('.', ',')

    hours = str(hours).replace('.', ',')
    issue_date = datetime.now().strftime('%d.%m.%Y')
    duedate = (datetime.now() + timedelta(days=14)).strftime('%d.%m.%Y')

    sheet = new_invoice.sheet1
    sheet.update('A1', f'Faktura č.{invoice_num}')
    sheet.update('B21', str(issue_date))
    sheet.update('B22', str(duedate))
    sheet.update('F21', invoice_num)
    sheet.update('E27', hours)
    sheet.update('G27', f'{total} Kč')
    sheet.update('G30', f'{total} Kč')

    access_token = creds.get_access_token().access_token

    url = ('https://docs.google.com/spreadsheets/d/' + new_invoice.id + '/export?'
           + 'format=pdf'  # export as PDF
           + '&gridlines=false'  # no gridlines
           + '&horizontal_alignment=CENTER'  # aligned horizontaly
           + '&vertical_alignment=CENTER'  # aligned verticaly
           + '&scale=1'  # 1=100%
           + '&access_token=' + access_token)  # access token

    request = requests.get(url)
    filename = f'{invoice_num}.pdf'
    with open(filename, 'wb') as saveFile:
        saveFile.write(request.content)


prepare_invoice(invoice_name)


if new_report in reports:
    logging.warning(f'Spreadsheet {new_report} for this month already exists')
    exit()
else:
    make_new_report(client, new_report)

import logging
from pathlib import Path
import re
import requests
import months_cz
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta


def rename_form(service, new_report_name):
    forms = service.files().list(
        q='mimeType=\'application/vnd.google-apps.form\' and name=\'Copy of template-report-form-naturaservis\'',
        fields='files(id)'
    ).execute()
    form_id = forms['files'][0]['id']
    logging.info(f'Renaming form to correct name: {new_report_name} [REMOTE]')
    service.files().update(fileId=form_id, body={'name': new_report_name}).execute()


def get_all_sheets(service):
    all_sheets = service.files().list(
        spaces='drive',
        q='mimeType=\'application/vnd.google-apps.spreadsheet\' and parents in \'1WIifQderN4MQ5gqyDtj3L6svIj7S6aHk\'',
        fields='files(id, name)').execute()

    sheets_dict = {}
    for s in all_sheets['files']:
        sheets_dict[s['name']] = s['id']

    return sheets_dict


# returns a tuple with name and id of the latest report
def get_latest_report(files):
    logging.info('Searching for latest report [REMOTE]')
    latest, latest_id, latest_name = 0, 0, 0

    for f in files.keys():
        if 'pracovni vykaz' in str(f):
            number = int(f[:2])
            if number > latest:
                latest = number
                latest_name = f
                latest_id = files[f]

    logging.info(f'Latest report found {latest_name}')
    return latest_name, latest_id


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
    # don't copy if report for this month exists on gdrive
    logging.info(f'Creating new report {report_name} [REMOTE]')
    template_report_id = '1Cpexu6iBaKRvDCO7uBV6dQr0-t3qfQQgPUDipGJvBcE'
    report = authorized_client.copy(template_report_id, title=report_name, copy_permissions=True)
    year = datetime.now().year
    report.sheet1.update('b4', f'měsíc: {get_month_name()[1].capitalize()} {year}')
    logging.info(f'Report created {report_name} [REMOTE]')


def export_last_report(file, service):

    file_name, file_id = file
    logging.info('Checking for duplicite report [LOCAL]')
    if not Path(f'reports/{file_name}.xlsx').exists():
        mime_type = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        byte_data = service.files().export(
            fileId=file_id,
            mimeType=mime_type,
        ).execute()

        logging.info(f'Exporting report {file_name}.xlsx')
        with open(f'reports/{file_name}.xlsx', 'wb') as f:
            f.write(byte_data)


def get_hours(report):
    return report.sheet1.acell('G6').value


def get_latest_invoice(all_sheets):

    last, last_name = 0, 0
    sheet_name_regex = re.compile(r'\d{8}')
    for sheet_name in all_sheets:
        if sheet_name_regex.search(sheet_name):
            number = int(sheet_name[4:])
            if number > last:
                last = number
                last_name = sheet_name
    return last_name


# returns something like 20220002
def new_invoice_name(latest_invoice):
    year = datetime.now().year
    invoice_num = int(latest_invoice[4:]) + 1
    new_spreadsheet_name = str(year) + str(invoice_num).zfill(4)
    return new_spreadsheet_name


def check_loger(client, latest_invoice):
    # loger = Path('loger.log')
    # if not loger.exists():
    #     loger.touch()

    last_invoice = client.open(latest_invoice).sheet1
    # Last 8 digits in A1; invoice num
    invoice_num = str(last_invoice.acell('A1').value)[-8:]
    issued = last_invoice.acell('B21').value
    due_date = last_invoice.acell('B22').value
    hours = last_invoice.acell('E27').value
    total = str(last_invoice.acell('G27').value)[:-3]
    return [invoice_num, issued, due_date, hours, total]

    # with open(loger, 'w') as log_file:
    #     pass


def prepare_invoice(client, report, latest_invoice):
    hours = get_hours(client.open(report))

    logging.info('Checking for existing invoice documents [REMOTE]')

    if latest_invoice == 0:
        invoice_num = str(datetime.now().year) + '0001'
        logging.info(f'No invoice document found [REMOTE]')
        previous_invoice_data = []
    else:
        invoice_num = new_invoice_name(latest_invoice)
        logging.info(f'New invoice document {invoice_num} will be created [REMOTE]')
        previous_invoice_data = check_loger(client, latest_invoice)

    # formating total cost to format 12 345,00
    total = 220 * float(hours)
    total = f'{total:,.2f}'
    total = total.replace(',', ' ').replace('.', ',')

    hours = str(hours).replace('.', ',')
    issue_date = datetime.now().strftime('%d.%m.%Y')
    # duedate = (datetime.now() + timedelta(months=14)).strftime('%d.%m.%Y')
    duedate = (datetime.now() + relativedelta(months=1)).strftime('%d.%m.%Y')
    current_invoice_data = [latest_invoice, issue_date, duedate, hours, total]

    logging.info(f'Checking for duplicite invoice document [REMOTE]')
    if not previous_invoice_data == current_invoice_data:
        logging.info(f'Creating new invoice with number {invoice_num} [REMOTE]')
        new_invoice = client.copy('1HJIUHBL2kRhS69OTKAfQSYRVSx0_3W1SwFLIdHaUylA', invoice_num, copy_permissions=True)

        sheet = new_invoice.sheet1
        logging.info('Populating data. [REMOTE]')
        sheet.update('A1', f'Faktura č.{invoice_num}')
        sheet.update('B21', str(issue_date))
        sheet.update('B22', str(duedate))
        sheet.update('F21', invoice_num)
        sheet.update('E27', hours)
        sheet.update('G27', f'{total} Kč')
        sheet.update('G30', f'{total} Kč')
        logging.info('Done populating [REMOTE]')

    else:
        logging.warning(f'New invoice with number {invoice_num} would have same exact data as {latest_invoice} [REMOTE]')


def export_latest_invoice(client, access_token, service):

    all_sheets = get_all_sheets(service)
    invoice_num = get_latest_invoice(all_sheets)
    invoices_dir = Path('invoices')
    filename = invoices_dir / f'{invoice_num}.pdf'

    if not filename.exists():
        logging.info(f'Exporting invoice {filename.name} [LOCAL]')
        new_invoice = client.open(invoice_num)
        url = ('https://docs.google.com/spreadsheets/d/' + new_invoice.id + '/export?'
               + 'format=pdf'  # export as PDF
               + '&gridlines=false'  # no gridlines
               + '&horizontal_alignment=CENTER'  # aligned horizontaly
               + '&vertical_alignment=CENTER'  # aligned verticaly
               + '&scale=1'  # 1=100%
               + '&access_token=' + access_token)  # access token

        request = requests.get(url)

        with open(filename, 'wb') as saveFile:
            saveFile.write(request.content)

        logging.info('Invoice export done [LOCAL]')
    else:
        logging.warning(f'Invoice {filename.name} already exists [LOCAL]')

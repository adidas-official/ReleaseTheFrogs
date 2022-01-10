# TOOD:
# 1) run daily to refresh token
#   - use Google computing
# 2) if end of month, download the latest report and make new spreadsheet                   [OK]
# 3) find date and hours in invoice document, change them and save as new invoice pdf       [OK]
# 4) make email with two attachements, report and invoice, subject, body and footer
# 5) send to blanka

from datetime import datetime, timedelta
from pathlib import Path
import ezsheets
import months_cz
import logging
import re
import downloader

logging.basicConfig(level=logging.DEBUG,
                    format='%(asctime)s - %(levelname)s - %(message)s',
                    datefmt='%Y-%m-%d %H:%M:%S')


def get_latest_invoice(all_sheets):

    last = 0
    sheet_name_regex = re.compile(r'\d{8}')
    for sheet_name in all_sheets.values():
        if sheet_name_regex.search(sheet_name):
            number = int(sheet_name[4:])
            if number > last:
                last = number
    return last


# finds the latest report
# returns sheet id
def get_latest_report(all_sheets):
    latest = 0
    latest_id = 0

    for sheet_id, sheet_name in all_sheets.items():
        if 'pracovni vykaz' in sheet_name:
            number = int(sheet_name[:2])
            if number > latest:
                latest = number
                latest_id = sheet_id

    return latest_id


def get_hours(report_id):
    ss = ezsheets.Spreadsheet(report_id)
    sheet = ss[0]
    hours = sheet['G6']
    return hours


def export_to_dest(report_id, dest):
    # dest = Path('C:/Users/bluem/Documents/Prace/NaturaServis/pracovni_vykaz/2022')
    ss = ezsheets.Spreadsheet(report_id)
    filename = dest / (ss.title + '.xlsx')

    if not filename.exists():
        logging.info(f'Exporting {filename.name}')
        ss.downloadAsExcel(dest / (ss.title + '.xlsx'))
    else:
        logging.info(f'{filename.name} already exists')


def start_of_month(dt):
    todays_month = dt.month
    yesterdays_month = (dt + timedelta(days=-1)).month
    return True if yesterdays_month != todays_month else False


# get name of current month
# use it for creating name of new spreadsheet
# returns number[string with leading zero] of month, and it's name
def get_month_name():
    month_now = datetime.now().month
    month_num = month_now % 12
    month_name = months_cz.months_cz[month_num - 1]
    month_num = str(month_num).zfill(2)
    return month_num, month_name


# get new name for template
# returns something like 01-pracovni vykaz leden, depending on current date
def new_report_name():
    month_num, month_name = get_month_name()
    new_spreadsheet_name = f'{month_num}-pracovni vykaz {month_name}'
    return new_spreadsheet_name


# returns something like 20220002
def new_invoice_name(latest_invoice):
    year = datetime.now().year
    new_spreadsheet_name = str(year) + str(int(latest_invoice + 1)).zfill(4)
    return new_spreadsheet_name


# fill b4 cell with correct month name
def fill_month(sheet, month):
    year = datetime.now().year

    if datetime.now().month == 12:
        year += 1

    updated_month = f'Měsíc: {month.capitalize()} {year}'
    sheet['B4'] = updated_month
    sheet['G6'] = '=SUM(F8:F38)'


def prepare_invoice(invoice_num, hours):
    ss = ezsheets.createSpreadsheet(invoice_num)
    invoice_template = ezsheets.Spreadsheet('1qDU1K8uempWTMXC8baHjlO-6EANEn2M7F1KhD4ziUMY')
    invoice_template[0].copyTo(ss)

    # remove first sheet, rename other to 'List1'
    ss[0].delete()
    ss[0].title = 'List1'

    # formating total cost to format 12 345,00
    total = 150 * float(hours)
    total = f'{total:,.2f}'
    total = total.replace(',', ' ').replace('.', ',')

    hours = str(hours).replace('.', ',')
    issue_date = datetime.now().strftime('%d.%m.%Y')
    duedate = (datetime.now() + timedelta(days=14)).strftime('%d.%m.%Y')

    sheet = ss[0]
    sheet['A1'] = f'Faktura č.{invoice_num}'
    sheet['B21'] = str(issue_date)
    sheet['B22'] = str(duedate)
    sheet['F21'] = invoice_num
    sheet['E27'] = hours
    sheet['G27'] = f'{total} Kč'
    sheet['G30'] = f'{total} Kč'

    ss.downloadAsPDF(Path(f'C:/Users/bluem/Documents/Prace/NaturaServis/faktury/{datetime.now().year}/{invoice_num}.pdf'))
    downloader.get_that_b(ss.spreadsheetId, invoice_num)


def main():

    all_sheets = ezsheets.listSpreadsheets()

    latest_invoice = get_latest_invoice(all_sheets)
    invoice_num = new_invoice_name(latest_invoice)

    latest_report = get_latest_report(all_sheets)
    hours = get_hours(latest_report)

    if invoice_num in all_sheets.values():
        logging.warning(f'Spreadsheet {invoice_num} for this month already exists')
        exit()
    else:
        prepare_invoice(invoice_num, hours)

    if start_of_month(datetime.now()):

        # download the latest report to local machine
        latest_report = get_latest_report(all_sheets)
        export_to_dest(latest_report, Path('C:/Users/bluem/Documents/Prace/NaturaServis/pracovni_vykaz/2022'))

        # get hours
        latest_report = get_latest_report(all_sheets)
        hours = get_hours(latest_report)

        # make spreadsheet for new month from template
        report_template = ezsheets.Spreadsheet('1AeZBElTTboZYDbNABKJQHtZlUi7rkQ-zgtjc5TXYc_Q')
        new_sheet_name = new_report_name()

        # check if report for current month already exists
        if new_sheet_name in all_sheets.values():
            logging.warning(f'Spreadsheet {new_sheet_name} for this month already exists')
            exit()
        else:
            report = ezsheets.createSpreadsheet(new_report_name())
            report_template[0].copyTo(report)

            # remove first sheet, rename other to 'List1'
            report[0].delete()
            report[0].title = 'List1'
            fill_month(report[0], get_month_name()[1])


if __name__ == '__main__':
    main()

import requests
import openpyxl
import datetime
from sys import argv

# Check of EfZ-Nachweis via DPSG NaMi 2.2 web interface
# data from a Microsoft Excel spreadsheet is processed and updated


def check_status(input_data):
    url = "https://nami.dpsg.de/ica/sgb-acht-bescheinigung-pruefen"
    msg_fail = '<p class="failure-msg">Bescheinigung ist NICHT gültig</p>'
    msg_success = '<p class="success-msg">Bescheinigung ist gültig</p>'

    # these could be checked beforehand querying the webpage!
    msg_invalid_id = '<p class="validation-msg">Identifikationsnummer: Feld ist ein Pflichtfeld.&lt;br&gt;</p>'
    msg_invalid_surname = '<p class="validation-msg">Nachname: Feld ist ein Pflichtfeld.&lt;br&gt;</p>'
    msg_invalid_firstname = '<p class="validation-msg">Vorname: Feld ist ein Pflichtfeld.&lt;br&gt;</p>'

    r = requests.post(url, input_data, timeout=5)
    status = "!!! UNBEKANNT !!!"

    if r.status_code == requests.codes.ok:
        if msg_success in r.text:
            status = "GÜLTIG"
        elif msg_fail in r.text:
            status = "!!! UNGÜLTIG !!!"
        elif msg_invalid_id in r.text:
            status = "!!! FZ-NR PFLICHTFELD !!!"
        elif msg_invalid_surname in r.text:
            status = "!!! NACHNAME PFLICHTFELD !!!"
        elif msg_invalid_firstname in r.text:
            status = "!!! VORNAME PFLICHTFELD !!!"
        else:
            # should not happen ;)

            # for debugging:
            print('--------------')
            print(r.text)
            print('--------------')
            pass
    return status


def process_excel_file(filename, update=True):
    wb = openpyxl.load_workbook(filename)
    sheet = wb.active

    # dependent on the specific file structure the data is spread across different
    # columns; we need to define the column indices here
    # (unless there's some auto-detection feature implemented, or a specific
    # default naming scheme is used)
    col_firstname = 1
    col_surname = 2
    col_fznumber = 3
    col_birthdate = 4
    col_status = 5
    col_update_timestamp = 6
    for i in range(2, sheet.max_row + 1):
        birthdate = sheet.cell(row=i, column=col_birthdate).value
        if birthdate:
            if type(birthdate) == datetime.datetime:
                birthdate = birthdate.strftime("%d.%m.%Y")
        else:
            birthdate = ""

        fznumber = sheet.cell(row=i, column=col_fznumber).value
        if not fznumber:
            fznumber = ""

        firstname = sheet.cell(row=i, column=col_firstname).value
        if not firstname:
            firstname = ""

        surname = sheet.cell(row=i, column=col_surname).value
        if not surname:
            surname = ""

        # TODO/FIXME: handle empty string
        in_data = {
            'fzNummer': fznumber,
            'vorname': firstname,
            'nachname': surname,
            'geburtsdatum': birthdate
        }
        # for debugging:
        # print(in_data)

        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        status = check_status(in_data)
        output_format = f"[{timestamp}] {in_data.get('vorname'):>12} {in_data.get('nachname'):>12}:  {status}"
        print(output_format)

        sheet.cell(row=i, column=col_status).value = status
        sheet.cell(row=i, column=col_update_timestamp).value = timestamp

    # data may be fed back to Excel into two columns
    if update:
        wb.save(filename)


if __name__ == "__main__":
    if len(argv) >= 2:
        process_excel_file(argv[1])
    else:
        print(f"Usage: python {argv[0]} <excel-filename>")

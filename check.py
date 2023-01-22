import requests
import openpyxl
import datetime
from sys import argv

# Check of EfZ-Nachweis via DPSG NaMi 2.2 web interface
# data from a Microsoft Excel spreadsheet is processed and updated


def query_efz_status(input_data):
    url = "https://nami.dpsg.de/ica/sgb-acht-bescheinigung-pruefen"
    msg_fail = '<p class="failure-msg">Bescheinigung ist NICHT gültig</p>'
    msg_not_valid = '<p class="failure-msg">Die Echtheit der Bescheinigung kann nicht bestätigt werden.</p>'
    msg_success = '<p class="success-msg">Das Dokument ist echt. Die Bescheinigung wurde korrekt erstellt.</p>'

    # these could be logically checked beforehand querying the webpage!
    msg_invalid_id = '<p class="validation-msg">Identifikationsnummer: Feld ist ein Pflichtfeld.&lt;br&gt;</p>'
    msg_invalid_surname = '<p class="validation-msg">Nachname: Feld ist ein Pflichtfeld.&lt;br&gt;</p>'
    msg_invalid_firstname = '<p class="validation-msg">Vorname: Feld ist ein Pflichtfeld.&lt;br&gt;</p>'
    msg_invalid_birthdate = '<p class="validation-msg">Geburtsdatum: Feld ist ein Pflichtfeld.&lt;br&gt;</p>'

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
        elif msg_invalid_birthdate in r.text:
            status = "!!! GEB.DATUM PFLICHTFELD !!!"
        elif msg_not_valid in r.text:
            status = "!!! NICHT BESTÄTIGT !!!"
        else:
            # should not happen ;)

            # print response for debugging:
            print('--------------')
            print(r.text)
            print('--------------')
            pass

        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    return status, timestamp



def get_query_input_from_sheet_row(sheet, row_idx):
    # dependent on the specific file structure the data is spread across different columns;
    # we need to define the column indices here;
    # we need first name, surname, EfZ number and birthday to start the query
    # TODO: make this a command line parameter (or somehow add auto-detection)
    #       (unless there's some auto-detection feature implemented, or a specific
    #       default naming scheme is used)
    col_firstname = 1
    col_surname = 2
    col_fznumber = 3
    col_birthdate = 4

    birthdate = sheet.cell(row=row_idx, column=col_birthdate).value
    if birthdate:
        if type(birthdate) == datetime.datetime:
            birthdate = birthdate.strftime("%d.%m.%Y")
    else:
        birthdate = ""

    fznumber = sheet.cell(row=row_idx, column=col_fznumber).value
    if not fznumber:
        fznumber = ""

    firstname = sheet.cell(row=row_idx, column=col_firstname).value
    if not firstname:
        firstname = ""

    surname = sheet.cell(row=row_idx, column=col_surname).value
    if not surname:
        surname = ""

    return {
        'fzNummer': fznumber,
        'vorname': firstname,
        'nachname': surname,
        'geburtsdatum': birthdate
    }


def process_excel_file(filename, update=True):
    # TODO: we could also make this CSV input format compatible
    wb = openpyxl.load_workbook(filename)
    sheet = wb.active

    # dependent on the specific file structure the data is spread across different columns;
    # we need to define the column indices here;
    # status and update timestamp are used to store the response in the same document
    # TODO: make this a command line parameter (or somehow add auto-detection)
    #       (unless there's some auto-detection feature implemented, or a specific
    #       default naming scheme is used)
    col_status = 5
    col_update_timestamp = 6

    start_row = 2  # skip the header row as it usually contains the headings
    # TODO: make this a command line parameter (or somehow add auto-detection)
    for row_idx in range(start_row, sheet.max_row + 1):
        in_data = get_query_input_from_sheet_row(sheet, row_idx=row_idx)

        efz_status, timestamp = query_efz_status(in_data)

        # visualize result (along with some input data) on the command line
        # TODO: add command line flag (verbosity? default=False)
        output_format = f"[{timestamp}] {in_data.get('vorname'):>20} {in_data.get('nachname'):>20}:  {efz_status}"
        print(output_format)

        # write result back to the workbook
        # TODO: add command line flag (name could be "--update"/"--dry-run"/"--print-only"; default=True)
        sheet.cell(row=row_idx, column=col_status).value = efz_status
        sheet.cell(row=row_idx, column=col_update_timestamp).value = timestamp

    # data may be fed back to Excel into two columns
    if update:
        wb.save(filename)


if __name__ == "__main__":
    if len(argv) >= 2:
        filename = argv[1]
        try:
            process_excel_file(filename)
        except FileNotFoundError:
            print(f"Could not find file '{filename}'.")
        except openpyxl.utils.exceptions.InvalidFileException:
            print(f"File type of '{filename}' is currently not supported (supported formats are .xlsx, .xlsm, .xltx and .xltm).")
    else:
        print(f"Usage: python {argv[0]} <excel-filename>")

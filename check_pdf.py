#!/usr/bin/env python3
import argparse
import requests
import datetime
from pypdf import PdfReader

QUERY_URL = "https://nami.dpsg.de/ica/sgb-acht-bescheinigung-pruefen"

# only use the following two to check failure and success
MSG_SUCCESS = '<p class="success-msg">'
MSG_FAILURE = '<p class="failure-msg">'


# TODO/FIXME: don't repeat yourself (DRY); this code has been copied from `check.py` (consider this being a "quick hack")
def query_efz_status(input_data):
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    complete = all(val != "" for val in input_data.values())

    if not complete:
        status = "Eingangsdaten unvollständig"  # TODO: could be more precise about what is missing
    else:
        status = "Ungültig oder Abfrage fehlerhaft"

        r = requests.post(QUERY_URL, input_data, timeout=5)

        try:
            if r.status_code == requests.codes.ok:
                if (MSG_SUCCESS in r.text) and not (MSG_FAILURE in r.text):
                    status = "Gültig"
        except:
            # swallow bad requests for robustness (so that the script won't stop for a single entry)
            # TODO: narrow it down and don't keep a bare exception handler this broad!
            pass

    return status, timestamp


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='%(prog)s checks single EfZ-Nachweis (PDF) in DPSG NaMi 2.2 using '
                                                 'HTTP requests')

    parser.add_argument(dest='filename',
                        help='path to PDF input file')

    args = parser.parse_args()

    reader = PdfReader(args.filename)
    text = reader.pages[0].extract_text()

    assert QUERY_URL in text, "Could not find the QUERY_URL in the PDF's text. Unexpected format."

    lines = text.splitlines()

    # for debugging only:
    # print(lines)

    expected_text = [
        '',
        'Bescheinigung über die Einsichtnahme in das erweiterte ',
        'Führungszeugnis',
        'Hiermit bestätigen wir, dass',
        'geboren am  ',
        'wohnhaft in ',
        'am   uns ein erweitertes Führungszeugnis vorgelegt hat. Das  ',
        'Führungszeugnis  mit  dem  Datum  vom   wurde  durch  uns  ',
        'eingesehen und enthielt im Sinne des §72a SGB VIII keine Eintragungen.',
        'Diese  Bestätigung  wurde  automatisch  generiert  und  ist  auch  ohne  ',
        'Unterschrift gültig. Die Echtheit dieses Dokument kann auf der folgenden  ',
        'Webseite geprüft werden.',
        'Identifikationsnummer: ',
        ' Fon: ',
        ' Fax: ',
        ' www.dpsg.de',
        ' Rechtsträger: Bundesamt Sankt Georg e.V.'
    ]

    # sanity check: check if expected text matches (ignoring whitespaces in front and back)
    for i in range(0, len(expected_text)):
        assert expected_text[i].strip() == lines[i].strip(), \
            "Expected PDF content format does not match. Format may have changed."
        # when this assertion breaks, please open an issue on GitHub!

    # for debugging only:
    # for i in range(len(expected_text), len(lines)):
    #    print(lines[i])

    offset = len(expected_text)
    fznumber = lines[offset + 2]
    name = lines[offset + 4].split(' ')
    surname = name[-1]  # assume that the last name does not contain a space and equals the last element
    firstname = ' '.join(name[:-1])  # all other names go into firstname
    birthdate = lines[offset + 9]

    request_data = {
        'fzNummer': fznumber,
        'vorname': firstname,
        'nachname': surname,
        'geburtsdatum': birthdate
    }

    # For debugging only:
    # print(request_data)

    efz_status, timestamp = query_efz_status(request_data)

    output_format = f"[{timestamp}] {request_data.get('vorname')} " \
                    f"{request_data.get('nachname')}   {efz_status}"
    print(output_format)

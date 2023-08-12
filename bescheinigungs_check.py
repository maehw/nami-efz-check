#!/usr/bin/env python3

class BescheinigungsCheck:
    QUERY_URL = "https://nami.dpsg.de/ica/sgb-acht-bescheinigung-pruefen"

    # only use the following two to check failure and success
    MSG_SUCCESS = '<p class="success-msg">'
    MSG_FAILURE = '<p class="failure-msg">'

    # these message have been seen in the HTTP response bodies (HTML)
    # '<p class="failure-msg">Bescheinigung ist NICHT gültig</p>'
    # '<p class="failure-msg">Die Echtheit der Bescheinigung kann nicht bestätigt werden.</p>'
    # '<p class="success-msg">Das Dokument ist echt. Die Bescheinigung wurde korrekt erstellt.</p>'

    # these are logically checked before querying the webpage and should never be seen!
    # '<p class="validation-msg">Identifikationsnummer: Feld ist ein Pflichtfeld.&lt;br&gt;</p>'
    # '<p class="validation-msg">Nachname: Feld ist ein Pflichtfeld.&lt;br&gt;</p>'
    # '<p class="validation-msg">Vorname: Feld ist ein Pflichtfeld.&lt;br&gt;</p>'
    # '<p class="validation-msg">Geburtsdatum: Feld ist ein Pflichtfeld.&lt;br&gt;</p>'

    def __init__(self, firstnames, surname, birthdate, number):
        self.firstnames = firstnames
        self.surname = surname
        self.birthdate = birthdate
        self.number = number
        self.check_input_data()
        self.check_status = self.check()

    def get_firstnames(self):
        return self.firstnames

    def get_surname(self):
        return self.surname

    def get_birthdate(self):
        return self.birthdate

    def get_id(self):
        return self.number

    def get_check_status(self):
        return self.check_status

    def check(self, timeout=5):
        import requests

        r = requests.post(BescheinigungsCheck.QUERY_URL,
                          {
                            'fzNummer': self.number,
                            'vorname': self.firstnames,
                            'nachname': self.surname,
                            'geburtsdatum': self.birthdate
                          },
                          timeout=timeout)

        if r.status_code == requests.codes.ok:
            if (BescheinigungsCheck.MSG_SUCCESS in r.text) and not (BescheinigungsCheck.MSG_FAILURE in r.text):
                return True
        return False

    def __repr__(self):
        return f"Vorname(n): {self.firstnames}\n" \
               f"Nachname: {self.surname}\n" \
               f"Geb.datum: {self.birthdate}\n" \
               f"ID: {self.number}\n" \
               f"Check-Status: {self.check_status}"

    def check_input_data(self):
        import re
        if not self.firstnames:
            raise ValueError('Invalid first name(s)')
        if not self.surname:
            raise ValueError('Invalid surname')

        if not self.birthdate:
            raise ValueError('Invalid birthdate')
        elif not re.match(r"\d{2}\.\d{2}\.\d{4}", self.birthdate):
            raise ValueError(f"Unexpected format of birthdate (got: '{self.birthdate}')")

        if not self.number:
            raise ValueError('Invalid ID')
        elif not re.match(r"\d{5}-\d{4,5}", self.number):
            raise ValueError(f"Unexpected format of ID (got: '{self.number}')")

    @classmethod
    def from_pdf_file(cls, filename):
        from pypdf import PdfReader

        reader = PdfReader(filename)
        text = reader.pages[0].extract_text()
        lines = text.splitlines()

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
            if expected_text[i].strip() != lines[i].strip():
                raise ValueError('Expected PDF content format does not match. Format may have changed.')
                # when this exception is thrown, please open an issue on GitHub!

        # extract the interesting parts:
        offset = len(expected_text)
        number = lines[offset + 2]
        name = lines[offset + 4].split(' ')
        surname = name[-1]  # assume that the last name does not contain a space and equals the last element
        firstnames = ' '.join(name[:-1])  # all other names go into firstname
        birthdate = lines[offset + 9]

        return cls(firstnames, surname, birthdate, number)

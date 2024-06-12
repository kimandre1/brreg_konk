import requests
from bs4 import BeautifulSoup
import pandas as pd
import re

#### Møkkaprogramm som scraper brreg.no sorry not sorry andreas ####

# Henr orgnr fra fil
excel_file = 'orgnummer.xlsx'  # bruk faktisk navn og sti
org_numbers_df = pd.read_excel(excel_file, usecols="A")
org_numbers = org_numbers_df.iloc[:, 0].tolist()

# Base URL for kunngjøringer
base_url = "https://w2.brreg.no/kunngjoring/hent_nr.jsp?orgnr="

# nøkkelord
status_keywords = [
    "slettet",
    "sletting",
    "avsluttet bobehandling",
    "innstilling av bobehandling",
    "konkursåpning",
    "sletting etter fusjon",
    "fusjonert",
    "tvangsoppløsning",
    "varsel om tvangsoppløsning",
    "fusjonsbeslutning",
]

def check_status(org_number):
    url = f"{base_url}{org_number}"
    response = requests.get(url)
    if response.status_code == 200:
        soup = BeautifulSoup(response.text, 'html.parser')
        text_lines = list(soup.stripped_strings)
        combined_text = ' '.join(text_lines).lower()

        for keyword in status_keywords:
            if keyword.lower() in combined_text:
                for i in range(len(text_lines) - 1):
                    line = text_lines[i].lower()
                    next_line = text_lines[i + 1].lower()

                    if re.match(r'\d{2}\.\d{2}\.\d{4}', line) and keyword in next_line:
                        date = line.strip()
                        keyword_line = next_line.strip()
                        return keyword, date, f"{date} {keyword_line}"
    return None, None, None

# Måke gjennom orgnummer og sjekke status
results = []
for org_number in org_numbers:
    status, date, line = check_status(org_number)
    if status:
        results.append((org_number, status, date, line))
        print(f"Organization number {org_number} har følgende status: {status} fra dato {date}. Linje: {line}")
    else:
        results.append((org_number, "Aktiv", "", ""))
        print(f"Orgnummer: {org_number} er aktiv")

# Lagre resultatet i ny excel dok
results_df = pd.DataFrame(results, columns=['Org Nummer', 'Status', 'Dato', 'Linje'])
results_df.to_excel('orgnr_status.xlsx', index=False)

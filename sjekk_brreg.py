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
        content = ' '.join(soup.stripped_strings).lower()
        for keyword in status_keywords:
            if keyword.lower() in content:
                # henter dato
                match = re.search(rf'{keyword.lower()}.*?(\d{{2}}.\d{{2}}.\d{{4}})', content)
                if match:
                    return keyword, match.group(1)
                else:
                    return keyword, "Ingen dato"
    return None, None

# måke gjennom liste med orgnummer
results = []
for org_number in org_numbers:
    status, date = check_status(org_number)
    if status:
        results.append((org_number, status, date))
        print(f"Organization number {org_number} har følgende status: {status} fra dato {date}")
    else:
        results.append((org_number, "Aktiv", ""))
        print(f"Orgnummer: {org_number} er aktiv")

# Lagre resultatet i ny excel dok
results_df = pd.DataFrame(results, columns=['Org Nummer', 'Status', 'Dato'])
results_df.to_excel('orgnr_status.xlsx', index=False)

##
# Purpose: Script to automate entering information from PDF into StudentAccess.
# Author: David Reynoso
# Date: 2019-07-15
##
import pandas as pd
from pdfminer.high_level import extract_text

def read_excel_data(excel_file):
    return pd.read_excel(excel_file)

def extract_pdf_data(pdf_files):
    pdf_data = []
    for pdf_file in pdf_files:
        text = extract_text(pdf_file)
        lines = text.split('\n')
        participant = {}

        gender_index = [i for i, line in enumerate(lines) if 'Gender' in line][0]
        if gender_index is not None and gender_index + 2 < len(lines):
            participant['gender'] = lines[gender_index + 2].strip()

        race_index = [i for i, line in enumerate(lines) if 'Race' in line][0]
        if race_index is not None and race_index + 2 < len(lines):
            participant['race'] = lines[race_index + 2].strip()

        birthday_index = [i for i, line in enumerate(lines) if 'Date of Birth' in line][0]
        if birthday_index is not None and birthday_index + 2 < len(lines):
            participant['birthday'] = lines[birthday_index + 2].strip()

        pdf_data.append(participant)

    return pdf_data

def process_data(df, pdf_data):
    excel_data = []
    for index, row in df.iterrows():
        participant_from_excel = {
            'first name': str(row['First Name']) if pd.notnull(row['First Name']) else '',
            'middle name': str(row['Middle Name']) if pd.notnull(row['Middle Name']) else '',
            'last name': str(row['Last Name']) if pd.notnull(row['Last Name']) else '',
            'bnumber': str(row['B Number']) if pd.notnull(row['B Number']) else ''
        }
        excel_data.append(participant_from_excel)

    for pdf_participant, excel_participant in zip(pdf_data, excel_data):
        print("Participant's Bnumber: " + excel_participant['bnumber'])
        full_name = " ".join([excel_participant['first name'], excel_participant['middle name'], excel_participant['last name']])
        print("Participant's Name: " + full_name)
        print("Participants Gender: " + pdf_participant['gender'])
        print("Participant's Race: " + pdf_participant['race'])
        print("Participant's Date of Birth: " + pdf_participant['birthday'])
        print("--------------------")

def main():
    excel_file = 'Sample Bnumber List.xlsx'
    df = read_excel_data(excel_file)

    pdf_files = ['Sample Data Entry David R.pdf', 'Sample Data Entry Nyx.pdf']
    pdf_data = extract_pdf_data(pdf_files)

    process_data(df, pdf_data)

if __name__ == "__main__":
    main()

##
# Purpose: Script to automate entering information from PDF and Excel file into StudentAccess.
# Author: David Reynoso
# Date: 2019-07-15
##
import pandas as pd
from pdfminer.high_level import extract_text
import glob
import re


def read_excel_data(excel_file):
    return pd.read_excel(excel_file)


def extract_pdf_data(pdf_files):
    pdf_data = []
    for pdf_file in pdf_files:

        text = extract_text(pdf_file)
        lines = text.split('\n')
        participant = {}

        # Extracting name from PDF
        name_indices = [i for i, line in enumerate(lines) if 'Name' in line]
        name_index = name_indices[0] if name_indices else None
        if name_index is not None and name_index + 2 < len(lines):
            participant['name'] = lines[name_index + 2].strip()

        gender_indices = [i for i, line in enumerate(lines) if 'Gender' in line]
        gender_index = gender_indices[0] if gender_indices else None
        if gender_index is not None and gender_index + 2 < len(lines):
            participant['gender'] = lines[gender_index + 2].strip()

        race_indices = [i for i, line in enumerate(lines) if 'Race' in line]
        race_index = race_indices[0] if race_indices else None
        if race_index is not None and race_index + 2 < len(lines):
            participant['race'] = lines[race_index + 2].strip()

        birthday_indices = [i for i, line in enumerate(lines) if 'Date of Birth' in line]
        birthday_index = birthday_indices[0] if birthday_indices else None
        if birthday_index is not None and birthday_index + 2 < len(lines):
            participant['birthday'] = lines[birthday_index + 2].strip()

        pdf_data.append(participant)
        # print(text) #TODO: Delete after use

    return pdf_data


def process_data(df, pdf_files):
    pdf_data = extract_pdf_data(pdf_files)
    excel_data = []
    student_info_list = []

    for index, row in df.iterrows():
        participant_from_excel = {
            'first name': str(row['First Name']) if pd.notnull(row['First Name']) else '',
            'middle name': str(row['Middle Name']) if pd.notnull(row['Middle Name']) else '',
            'last name': str(row['Last Name']) if pd.notnull(row['Last Name']) else '',
            'bnumber': str(row['B Number']) if pd.notnull(row['B Number']) else ''
        }
        excel_data.append(participant_from_excel)

    for pdf_participant, excel_participant in zip(pdf_data, excel_data):
        student_info = {
            'first name': excel_participant['first name'],
            'middle name': excel_participant['middle name'],
            'last name': excel_participant['last name'],
            'bnumber': excel_participant['bnumber'],
            'gender': pdf_participant['gender'],
            'Date of Birth': pdf_participant['birthday'],
            'Race': pdf_participant['race'],

        }
        student_info_list.append(student_info)

    return student_info_list


if __name__ == '__main__':  # the following function is a test function that only runs within the file itself.
    excel_file = 'Sample Bnumber List.xlsx'
    pdf_files = glob.glob("PDF Files/*.pdf")
    df = read_excel_data(excel_file)

    # Read all the PDF files
    pdf_data = extract_pdf_data(pdf_files)

    # Create a dictionary from Excel to easily lookup
    excel_dict = {}
    for index, row in df.iterrows():
        name = str(row['First Name']) + ' ' + str(row['Middle Name']) + ' ' + str(row['Last Name'])
        excel_dict[name] = {
            'first name': str(row['First Name']),
            'middle name': str(row['Middle Name']),
            'last name': str(row['Last Name']),
            'bnumber': str(row['B Number']),
        }

    # Match PDF and Excel based on name
    student_info_list = []
    for pdf_participant in pdf_data:
        if 'name' in pdf_participant:  # Confirm that the 'name' key exists
            name_from_pdf = pdf_participant['name'].strip()
            name_from_pdf = name_from_pdf.strip().lower()

            excel_participant = excel_dict.get(name_from_pdf, {})

            print(f"Name from PDF: {name_from_pdf}")
            print(f"Excel Participant: {excel_participant}")
            print(f"Excel Dict Sample: {list(excel_dict.keys())[:5]}")

            if not excel_participant:  # If name didn't match, let's look for the next line
                for i, (key, value) in enumerate(excel_dict.items()):
                    if name_from_pdf in value['first name'] + ' ' + value['last name']:
                        excel_participant = value
                        break

        if excel_participant:
            student_info = {
                'first name': excel_participant['first name'],
                'middle name': excel_participant['middle name'],
                'last name': excel_participant['last name'],
                'bnumber': excel_participant['bnumber'],
                'gender': pdf_participant['gender'],
                'Date of Birth': pdf_participant['birthday'],
                'Race': pdf_participant['race'],
            }
            student_info_list.append(student_info)

    print(student_info_list)

##
# Purpose: Script to automate entering information from PDF and Excel file into StudentAccess.
# Author: David Reynoso
# Date: 2019-07-15
##
import pandas as pd
from pdfminer.high_level import extract_text
import glob


def read_excel_data(excel_file):
    return pd.read_excel(excel_file)


# function that will create a dictonary from the excel data
def extract_excel_data(excel_file):
    df = read_excel_data(excel_file)
    excel_dict = {}
    for index, row in df.iterrows():
        information = (str(row['First Name'])).lower().strip()
        excel_dict[information] = {
            'first name': str(row['First Name']),
            'middle name': str(row['Middle Name']),
            'last name': str(row['Last Name']),
            'bnumber': str(row['B Number']),
        }
    return excel_dict
#print(extract_excel_data('Sample Bnumber List.xlsx'))


#(extract_excel_data('Sample Bnumber List.xlsx'))  # test print to see if all items are being printed.


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
            full_name = lines[name_index + 2].strip()
            first_name = full_name.split(' ')[0]  # only gets the first word of the name
            participant['name'] = first_name.lower().strip()

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
#print(extract_pdf_data(glob.glob("PDF Files/*.pdf")))







# function that will cross-reference the list of dictionaries from the PDFs with the Excel dictionary.
def match_data(pdf_data, excel_dict):
    combined_data = []
    for pdf_participant in pdf_data:
        # tries to find a match in the Excel dictionary
        excel_participant = excel_dict.get(pdf_participant.get('name', '').lower())
        if excel_participant:
            #combines dictionaries
            combined_info = {**excel_participant, **pdf_participant}
            combined_data.append(combined_info)
    print(combined_data)
    return combined_data


'''def process_data(df, pdf_files):
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
'''  # function that will create a list of dictionaries from the excel and pdf data

if __name__ == "__main__":
    excel_data = extract_excel_data('Sample Bnumber List.xlsx')
    pdf_data = extract_pdf_data(glob.glob("PDF Files/*.pdf"))
    combined_data = match_data(pdf_data, excel_data)
    print("Combined Data: ", combined_data)

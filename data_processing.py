##
# Purpose: Script to automate pulling information from the excel file and PDF's and storing it in a dictionary.
# Author: David Reynoso
# Date: 2019-07-15
##
import pandas as pd
from pdfminer.high_level import extract_text
import glob


def read_excel_data(excel_file):
    excel_data = pd.read_excel(excel_file)  # Read the excel data into a DataFrame
    excel_dict = {}  # Initialize an empty dictionary

    for index, row in excel_data.iterrows():
        composite_key = f"{row['First Name'].strip().lower()} {row['Middle Name'].strip().lower()} {row['Last Name'].strip().lower()}"
        excel_dict[composite_key] = {
            'B Number': row['B Number'],
            'first name': row['First Name'],
            'middle name': row['Middle Name'],
            'last name': row['Last Name'],
        }
    return excel_dict




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
            name_parts = full_name.split(' ')

            first_name = name_parts[0]
            last_name = name_parts[-1]
            if len(name_parts) > 2:
                middle_name = ' '.join(name_parts[1:-1])
            else:
                middle_name = '' #no middle name
            composite_key = f"{first_name.lower().strip()} {middle_name.lower().strip()} {last_name.lower().strip()}"
            participant['composite_key'] = composite_key

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
    print("Debugging: Type and value of pdf_data of extract_pdf_data: ", type(pdf_data), pdf_data)
    return pdf_data







# function that will cross-reference the list of dictionaries from the PDFs with the Excel dictionary.
def match_data(pdf_data, excel_dict):
    combined_data = []
    for pdf_participant in pdf_data:
        #print(f"Checking PDF participant: {pdf_participant.get('composite_key', '')}")
        # tries to find a match in the Excel dictionary
       # print("Debugging: Type and value of pdf_participant: ", type(pdf_participant), pdf_participant)
        #print("Checking PDF participant:", pdf_participant.get('composite_key', 'Key not found'))

        excel_participant = excel_dict.get(pdf_participant.get('composite_key', '').lower())
        if excel_participant:
            full_name_list = pdf_participant['composite_key'].split(' ')
            first_name = full_name_list[0]
            last_name = full_name_list[-1]
            #combines dictionaries
            student_info = {
                'first name': first_name,
                'middle name': excel_participant.get('middle name', ''),
                'last name': last_name,
                'B Number': excel_participant.get('B Number', ''),
                'gender': pdf_participant.get('gender', ''), # if changed to assigned gender at birth, change as well.
                'Date of Birth': pdf_participant.get('birthday',),
                'Race': pdf_participant.get('race',''),

            }

            combined_data.append(student_info)

    return combined_data




if __name__ == "__main__":
    excel_data = read_excel_data('Sample Bnumber List.xlsx')
    #print("Excel Data: ", excel_data)
    pdf_data = extract_pdf_data(glob.glob("PDF Files/*.pdf"))
   # print("PDF Data: ", pdf_data)
    combined_data = match_data(pdf_data, excel_data)
   # print("Combined Data: ", combined_data)

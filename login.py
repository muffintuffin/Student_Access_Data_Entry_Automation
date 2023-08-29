# Purpose: To login to the Berea SSS website navigate around the page
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from data_processing import process_data, read_excel_data, extract_pdf_data
from selenium.webdriver.support.ui import Select
import os

# function to log into website
def login_to_site(browser, username, password):
    # Opens webpage
    browser.get('https://berea-sss.studentaccess.com/login.aspx')

    # Finds username and password fields
    username_field = browser.find_element(By.ID, 'txtUsername')
    username_field.send_keys(username)
    password_field = browser.find_element(By.ID, 'txtPassword')
    password_field.send_keys(password)

    # Finds the login button and clicks it
    login_button = browser.find_element(By.ID, 'cmdLogin')
    login_button.click()


# navigates to the add student page
def navigate_to_add_student(browser):
    # Clicks on the add student link
    add_student = browser.find_element(By.LINK_TEXT, 'Add Student')
    add_student.click()


# enters in student information into text boxes
def step_one(browser, student_info):
    first_name_field = browser.find_element(By.ID, 'txtFirstName')
    first_name_field.send_keys(student_info['first name'])
    middle_name_field = browser.find_element(By.ID, 'txtMidName')
    middle_name_field.send_keys(student_info['middle name'])
    last_name_field = browser.find_element(By.ID, 'txtLastName')
    last_name_field.send_keys(student_info['last name'])
    bnumber_field = browser.find_element(By.ID, 'txtSID')
    bnumber_field.clear()
    bnumber_field.send_keys(student_info['bnumber'])
    save_and_edit_button = browser.find_element(By.ID, 'InsertButton')
    save_and_edit_button.click()


# enters in permanent information into text boxes
def enter_permanent_info(browser, student_info):
    DOB_field = browser.find_element(By.ID, 'txtDOB')
    DOB_field.send_keys(student_info['Date of Birth'])  # format: mm/dd/yyyy

    race_lower = student_info['Race'].lower()
    print(f"Lowercase race info: {race_lower}")  # Debug print

    # allows the program to select the participants race
    dropdown_hispanic = Select(browser.find_element(By.ID, 'ddlHispanic'))
    if 'hispanic' in race_lower or 'latino' in race_lower:
        dropdown_hispanic.select_by_value('Yes')
    else:
        dropdown_hispanic.select_by_value('No')

    dropdown_native = Select(browser.find_element(By.ID, 'ddlAmIndAK'))
    if 'american indian' in race_lower or 'alaskan native' in race_lower:
        dropdown_native.select_by_value('Yes')
    else:
        dropdown_native.select_by_value('No')

    dropdown_asian = Select(browser.find_element(By.ID, 'ddlAsian'))
    if 'asian' in race_lower:
        dropdown_asian.select_by_value('Yes')
    else:
        dropdown_asian.select_by_value('No')

    dropdown_black = Select(browser.find_element(By.ID, 'ddlBlackAfrAm'))
    if 'black/african american' in race_lower or 'African American' in race_lower:
        dropdown_black.select_by_value('Yes')
    else:
        dropdown_black.select_by_value('No')

    dropdown_white = Select(browser.find_element(By.ID, 'ddlWhite'))
    if 'white' in race_lower:
        dropdown_white.select_by_value('Yes')
    else:
        dropdown_white.select_by_value('No')

    dropdown_pacific = Select(browser.find_element(By.ID, 'ddlHIPacIslndr'))
    if 'native hawaiian' in race_lower or 'Other Pacific Islander' in race_lower:
        dropdown_pacific.select_by_value('Yes')
    else:
        dropdown_pacific.select_by_value('No')
    # selects the cohort, update it to match the current year's cohort.
    dropdown_cohort = Select(browser.find_element(By.ID, 'ddlCohort'))
    dropdown_cohort.select_by_value('2023-2024')

    dropdown_iel = Select(browser.find_element(By.ID, 'ddlInstEntryGradeLevel'))
    dropdown_iel.select_by_value('1st yr., never attended')
    # inserts in the date the student entered the program into Insitutional Entry Date, update as needed.
    ied_field = browser.find_element(By.ID, 'txtInstDate')
    ied_field.send_keys('08/23/2023')

    dropdown_iel = Select(browser.find_element(By.ID, 'ddlFirstServEnrollCD'))
    dropdown_iel.select_by_value('Full-time (at least 24 credit hours or 36 clock hours in an academic year)')

    # inserts in the date the student entered the program into First Service Date, update as needed.
    entry_date_field = browser.find_element(By.ID, 'txtEntryDate')
    entry_date_field.send_keys('08/21/2023')

    dropdown_last_service_date = Select(browser.find_element(By.ID, 'ddlAltLastServDate'))
    dropdown_last_service_date.select_by_value('88888888')

    dropdown_under_grad_date = Select(browser.find_element(By.ID, 'ddlAltUnderGradDate'))
    dropdown_under_grad_date.select_by_value('88888888')

    dropdown_degree_completed = Select(browser.find_element(By.ID, 'ddlDegreeCompleted'))
    dropdown_degree_completed.select_by_value('No degree/certificate, still enrolled at grantee institution')

    dropdown_certificate_earned = Select(browser.find_element(By.ID, 'ddlDegCertFieldErnd'))
    dropdown_certificate_earned.select_by_value('Has not earned a degree/certificate')

    #clikcs the save button
    save_button = browser.find_element(By.ID, 'cmdSave')
    save_button.click()


def main():
    # Inserts username and password
    username = "reynosod"
    password = "dpassword1!"

    # Opens up Browser and logs into website
    browser = webdriver.Firefox()
    login_to_site(browser, username, password)




    excel_file = 'Sample Bnumber List.xlsx'

    df = read_excel_data(excel_file)
    pdf_files = ['First Name Test.pdf', 'First Name Test 2.pdf', 'First Name Test 3.pdf']
    student_info_list = process_data(df, pdf_files)

    for student_info in student_info_list:
        navigate_to_add_student(browser)
        step_one(browser, student_info)
        enter_permanent_info(browser, student_info)



if __name__ == "__main__":
    main()

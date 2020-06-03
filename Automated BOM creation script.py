"""The purpose of this script is to automate the process of entering the information about the Bill of
Materials (BOM) in the AIMS portal. Running this script will lead to the auto-filling of the fields on
the BOM creation page.

The workflow of this script is as follows:
(1) Open a browser and go the AIMS web page
(2) Enter username and password and will click on the login button
(3) Click on the BOMs Tab in the portal
(4) Click on the add button
(5) Populate the fields on the BOM creation page
(6) Choose the BOM to be uploaded
(7) Save the BOM
(8) Repeat the above steps for other BOMs
"""

from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.keys import Keys
import xlrd
import os

# Enter the path to the folder where BOMs excel sheets have been kept
path_to_bom_folder = input("Enter the full path to the folder containing the BOMs: ")

browser = webdriver.Chrome()  # Opening the Chrome browser using the chromeDriver

# Extracting the BOM names from the folder containing the BOMs
list_bom_names = []

for file in os.listdir(path_to_bom_folder):
    if file.endswith(".xlsx"):
        list_bom_names.append(file)


def wait_until_element_becomes_clickable(x_path_of_element):

    """Script will use this function to wait until the web page loads completely or the element
    to be clicked becomes clickable. The function takes as its argument the Xpath of the element
    to be clicked"""

    # Script will wait for 0.5 sec for any web page to be loaded, after this time the script will declare
    # that the page took too long to load and will keep waiting for this page to be loaded
    delay = 1
    try:
        WebDriverWait(browser, delay).until(ec.element_to_be_clickable((By.XPATH, x_path_of_element)))
        print("Page is ready!")
    except TimeoutException:
        print("Loading took too much time!")


def login_to_portal():

    """Using the login_to_portal function, script will login to the AIMS portal."""

    url = "https://aims.adonmo.com"  # url for the AIMS portal
    browser.get(url)

    # Using the below function call, script will wait until the login button becomes clickable
    wait_until_element_becomes_clickable('//*[@id="login-form"]/input[2]')

    # Now, script will enter the username and password
    username = browser.find_element_by_name("username")
    password = browser.find_element_by_name("password")
    username.send_keys("varun@adonmo.com")
    password.send_keys("v@run.1234")

    # Finally, script will click on the login button on the web page
    login_attempt = browser.find_element_by_xpath('//*[@id="login-form"]/input[2]')
    login_attempt.submit()


def go_to_bom_tab():

    """Using this function, script will go to the BOMs tab in AIMS portal"""

    go_to_bom = browser.find_element_by_xpath('//*[@id="accordionSidebar"]/li[5]/a/span')
    go_to_bom.click()  # Script will click on the BOMs tab in the AIMS portal


def click_on_add_button():

    """Script will use this function to click on the add button on the BOMs web page"""

    add_bom = browser.find_element_by_xpath('//*[@id="content"]/div/div/div[1]/a[1]/i')
    add_bom.click()


def populate_fields():

    """Script will use this function to populate the following fields on the BOM creation page of the
    AIMS portal
    (1) PartNumber
    (2) Qty
    (3) Remarks"""

    # Entering the BOM name from the excel sheet to the "BOM Name" field
    bom_title = browser.find_element_by_xpath('//*[@id="id_bom_name"]')
    bom_title.send_keys(worksheet.cell(0, 0).value)

    # In case the excel sheets is empty then the remove button will take the value of i = 0 so that to
    # remove the very first row itself. To get better understanding, have a look at the next block of code
    # for the remove button

    i = 0
    for i in range(2, num_rows):

        # Entering the data in the search field
        search_button = browser.find_element_by_class_name('select2-selection__placeholder')
        search_button.click()
        search_field = browser.find_element_by_class_name('select2-search__field')
        search_field.send_keys(worksheet.cell(i, 1).value)

        # Script will wait until the entered element name appears in the drop down list
        wait_until_element_becomes_clickable('//*[@id="select2-bhv6-results"]')

        search_field.send_keys(Keys.ENTER)  # Clicking the option showing in the drop down list

        # Populating the qty field of the BOM creation page
        qty_field = browser.find_element_by_xpath('//*[@id="table_body"]/tr[{}]/td[4]/input'.format(i - 1))
        if worksheet.cell(i, 3).value == int(worksheet.cell(i, 3).value):
            qty_field.send_keys(int(worksheet.cell(i, 3).value), Keys.TAB)
        else:
            qty_field.send_keys(str(worksheet.cell(i, 3).value), Keys.TAB)

        # Populating the remarks field on BOM creation page
        remarks_field = browser.find_element_by_xpath('//*[@id="table_body"]/tr[{}]/td[5]/input'.format(i - 1))
        remarks_field.send_keys(worksheet.cell(i, 6).value)

        # Adding a new row after populating the previous row
        add_row_button = browser.find_element_by_xpath('//*[@id="add_row"]/span')
        add_row_button.click()

        # Script will wait until the next row appears
        wait_until_element_becomes_clickable('select2-selection__placeholder')

    # Removing an extra row
    remove_row_button = browser.find_element_by_xpath('//*[@id="table_body"]/tr[{}]/td[6]/a/span'.format(i))
    remove_row_button.click()


login_to_portal()  # Login to the AIMS portal

# Wait until the BOMs tab button shows
wait_until_element_becomes_clickable('//*[@id="accordionSidebar"]/li[5]/a/span')

go_to_bom_tab()  # Go to BOMs tab in the portal

bom_name = ""  # If the BOM folder is empty

# Script will now add each BOM to the portal, one at a time
for bom_name in list_bom_names:
    # Opening the excel sheet from which entries have to be imported
    workbook = xlrd.open_workbook("{}/{}".format(path_to_bom_folder, bom_name))
    worksheet = workbook.sheet_by_index(0)
    num_rows = worksheet.nrows

    # Script will wait until the add button becomes clickable
    wait_until_element_becomes_clickable('//*[@id="content"]/div/div/div[1]/a[1]')
    click_on_add_button()  # Going to BOM creation page

    # Wait until the BOM creation page loads completely
    wait_until_element_becomes_clickable('//*[@id="content"]/div/div/div[2]/form/a/span')

    populate_fields()  # Let's populate the relevant fields using excel sheet data :)

    # Uploading a file
    Upload_Button = browser.find_element_by_id('id_bom_file')
    Upload_Button.send_keys("{}/{}".format(path_to_bom_folder, bom_name))
    Upload_Button.submit()  # Saving the BOM after creation

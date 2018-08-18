"""
This is an automated way to upload an equipment list to CxAlloy.

There are several default options assumed. Ask JAD.
"""

from selenium import webdriver
import time
from selenium.webdriver.common.keys import Keys
import ctypes  # used for user input pause to select project
import pandas as pd
import difflib
import tkinter as tk
import sys

driver = webdriver.Chrome(
    "ADD PATH TO DRIVER HERE")

UN = sys.argv[1]
PW = sys.argv[2]

def site_login():
    """Navigate to login page, send login details."""
    driver.get("https://tq.cxalloy.com/auth/login")
    driver.find_element_by_xpath("//*[@id='username']").send_keys(UN)
    driver.find_element_by_xpath("//*[@id='password']").send_keys(PW)
    driver.find_element_by_id('login-button').click()


def select_project():
    """Selecting project and navigate to test templates."""
    ctypes.windll.user32.MessageBoxW(0, "Click CxAlloy project in the browser then click OK to continue.", "Navigate to CxAlloy Project Dashboard", 0x1000)
    url_project_num = (driver.current_url.rsplit('/', 1)[-1])

    driver.get("https://tq.cxalloy.com/project/" +
               url_project_num + "/import/equipment")


def select_files():
    """Get filepath of folder with excel docs from user."""
    import tkinter as tk
    from tkinter import filedialog
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(
        title="Select Equipment List File for Upload")
    return(file_path)


def path_to_name(file_path):
    """Take the file paths and return the file names."""
    import os
    return(os.path.splitext(os.path.basename(file_path))[0])


def read_excel(file_path):
    print("file path: " + file_path)
    xl = pd.ExcelFile(file_path)
    print("sheet name: " + xl.sheet_names[0])
    df = xl.parse(xl.sheet_names[0])
    print(len(df.columns))
    return(len(df.columns))


def name_id(header):
    """This function should identify the IMPORT AS selector based on the column header input."""
    pass
    if "Name" in header:
        return("Name")
    elif "Attribute" in header:
        return("Attribute")
    elif "Equipment Type" in header:
        return("Equipment Type")
    elif "Discipline" in header:
        return("Discipline")
    elif "Space" in header:
        return("Space")
    elif "System" in header:
        return("System")
    elif "Description" in header:
        return("Description")


def attribute_id(header, attribute_options):
    """INPUT:attribute_name_in and output:suggested attribute from list of avalible attributes."""
    attribute_name_in = (header[header.find("(") + 1:header.find(")")])
    print("______________________________")
    suggestions = difflib.get_close_matches(
        attribute_name_in, attribute_options, 1)  # get 1 closest matches
    print('attribute_name_in', '[suggestion]')
    print(attribute_name_in, suggestions)

    if len(suggestions) > 0:
        return(suggestions[0], None)
# IN PROGRESS BELOW##########################################!!!!!!!!!!!!!!!!!
    else:
        # handle edge case of no match better than this
        return("Tracking VAR", header)
# IN PROGRESS ABOVE##########################################!!!!!!!!!!!!!!!!!!!!!!!!!!!!


def fetch_attributes():
    """This will get a list of avalible attributes from the project."""
    html_list = driver.find_element_by_class_name(
        "mychzn-results")  # list of avalible attributes
    items = html_list.find_elements_by_tag_name("li")
    attribute_options = []
    for item in items:
        attribute_options.append(item.get_attribute('innerHTML'))
    if "" not in attribute_options:
        print(len(attribute_options), "fetch_attributes successfull")
    else:
        print("fetch_attributes FAILED!" + ("!" * 5))
        print(attribute_options)
    return(attribute_options)


def messageWindow(unmatched_items):
    """tkinter window with unmatched_items to check for attributes."""
    unmatched_items.insert(
        0, "Review matches and units for attributes then click OK to continue. The following were unable to be matched:\n")
    root = tk.Tk()
    root.title("Review The Following")
    label = tk.Label(root, text="\n".join(
        map(str, list(filter((None).__ne__, unmatched_items)))))
    label.pack(side="top", fill="both",
               expand=True, padx=20, pady=20)
    button = tk.Button(root, text="Continue", command=lambda: root.quit())
    button.pack(side="bottom", fill="none", expand=True)
    root.mainloop()


def add_templates(file_name, file_path, column_count):
    driver.find_element_by_class_name(
        "checkbox-label").click()  # clicking checkbox to update existing equipment
    driver.find_element_by_name("file").send_keys(file_path)
    driver.find_element_by_class_name("submit").click()  # continue
    time.sleep(1)
    driver.find_element_by_class_name("submit").click()  # continue
    time.sleep(2)  # may need to be 2

    unmatched_items = []
    for column in range(1, column_count + 1):
        # print("column: " + str(column))
        header = driver.find_element_by_xpath(
            "//*[@id='fields-form']/div[1]/div[2]/div[" + str(column) + "]/ul/h3")  # excel header
        import_as = name_id(header.text)

        driver.find_element_by_xpath(
            "//*[@id='fields-form']/div[1]/div[2]/div[" + str(column) + "]/div/select").send_keys(import_as)

        if import_as == "Attribute":
            # finds the column-######
            attribute_num = (driver.find_element_by_xpath(
                "//*[@id='fields-form']/div[1]/div[2]/div[" + str(column) + "]/div/select[1]").get_attribute("name")).split("-", 1)[1]
            # this is the attribute number for use in sending keys to cxalloy
            time.sleep(1)
            driver.find_element_by_xpath("//*[@id ='attribute_select_" +
                                         str(attribute_num) + "_chzn']/a").click()

            # set attribute_options when it doesn't exist
            try:
                attribute_options
            except NameError:
                attribute_options = fetch_attributes()

            # function to give match
            use_attribute, unmatched_item = attribute_id(
                header.text, attribute_options)
            unmatched_items.append(unmatched_item)
            time.sleep(.5)
            driver.find_element_by_xpath("//*[@id='attribute_select_" + str(  # send text to USE ATTRIBUTE box
                attribute_num) + "_chzn']/div/div/input").send_keys(use_attribute)
            time.sleep(.5)
            driver.find_element_by_xpath("//*[@id='attribute_select_" + str(  # send return to USE ATTRIBUTE BOX
                attribute_num) + "_chzn']/div/div/input").send_keys(Keys.RETURN)
    messageWindow(unmatched_items)

    time.sleep(.5)
    driver.find_element_by_class_name("submit").click()  # continue
    time.sleep(.5)
    # dont import the first row
    driver.find_element_by_xpath("//*[@id='rows']/tbody/tr[1]/td[1]/a").click()
    driver.find_element_by_class_name("submit").click()  # continue
    ctypes.windll.user32.MessageBoxW(
        0, "Automated process complete and successfull. Disregard CxAlloy Error Message at end. Please continue manually.", "Import Complete", 1)



site_login()
select_project()

file_path = select_files()  # list selected file path
file_name = path_to_name(file_path)  # list selected file name
column_count = read_excel(file_path)


add_templates(file_name, file_path, column_count)

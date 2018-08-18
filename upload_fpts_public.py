"""
This is an automated way to upload multiple FPT test templates to CxAlloy.

There are several default options assumed. Ask JAD.
"""

from selenium import webdriver
import time
from selenium.webdriver.common.keys import Keys
import ctypes  # used for user input pause to select project
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
    driver.find_element_by_xpath("//*[@id='test']/a").click()  # tests
    time.sleep(1)
    driver.find_element_by_xpath(
        "//*[@id='test-menu']/ul[1]/li[2]/a").click()  # test templates


def select_files():
    """Get filepath of folder with excel docs from user."""
    import tkinter as tk
    from tkinter import filedialog

    root = tk.Tk()
    root.withdraw()
    file_paths = list(filedialog.askopenfilenames(
        title="Select Functional Test Templates for Upload"))
    return(file_paths)


def path_to_name(file_paths):
    """Take the file paths and return the file names."""
    import os
    file_names = []
    for path in file_paths:
        file_names.append(os.path.splitext(os.path.basename(path))[0])
    return(file_names)


def add_templates(file_name, file_path):
    """Adding a new test template."""
    driver.find_element_by_xpath(
        "//*[@id='button-add-template']").click()  # add new
    time.sleep(1)
    # send title of test template
    driver.find_element_by_xpath(
        "//*[@id='template-name']").send_keys(file_name)
    # time.sleep(1)  # unneeded?
    driver.find_element_by_xpath(
        "//*[@id='button-submit-add-template']").click()  # add
    time.sleep(1)

    webelement = driver.find_element_by_xpath("//*[@id='project']/a")
    # sends 14 tabs and one return
    webelement.send_keys(Keys.TAB * 14 + Keys.RETURN)

    time.sleep(1)
    driver.find_element_by_class_name('import-button').click()  # import

    fancyboxFrame = driver.find_element_by_class_name(
        "fancybox-iframe")  # finding the fancybox popup
    # switching the focus of selenium to the fancybox
    driver.switch_to.frame(fancyboxFrame)
    # clicking checkbox for "background color to indicate headers"
    driver.find_element_by_class_name("checkbox-label").click()

    # this file path for upload and will need to be generated and changed in a loop
    driver.find_element_by_name("file").send_keys(file_path)

    driver.find_element_by_class_name("submit").click()  # continue
    driver.find_element_by_class_name("submit").click()  # continue
    driver.find_element_by_class_name(
        "import-select").send_keys("D")  # Import as description
    driver.find_element_by_class_name("submit").click()  # continue
    driver.find_element_by_class_name("submit").click()  # continue
    driver.find_element_by_class_name("submit").click()  # continue


site_login()
select_project()

file_paths = select_files()  # list of all selected file paths
file_names = path_to_name(file_paths)  # list of all selected file names

for name, path in zip(file_names, file_paths):  # add template for all files
    add_templates(name, path)

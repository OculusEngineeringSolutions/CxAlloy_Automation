# CxAlloy_Automation

Three different scripts are included in the repo for automation of typical user tasks using the online commissioning managment platform cxalloy.com.

# Bulk upload of functional test procedures
This script automates the login and navigation process to prepare for the checklist file upload using the Selenium library and the chrome web dirver. Next the user is prompted for multiple file selection using Tkinter. Once the user selects files the program will loop through the process for upload, sorting and naming for each file in turn again using Selenium for browser automation for all user selected files.

# Bulk upload of functional test procedures
This script automates the login and navigation process to prepare for the functional test file upload using the Selenium library and the chrome web driver. Next the user is prompted for multiple file selection using Tkinter. Once the user selects files the program will loop through the process for upload, sorting and naming for each file in turn again using Selenium for browser automation for all user selected files.

# Automated upload and sorting of equipment(asset) lists.
This script automates the login and navigation process to prepare for the equipment list (asset) file upload using the Selenium library and the chrome web driver. Next the user is prompted for a single file selection using Tkinter. Once the user selects the file including the equipment list the program will follow the process and automatically fill out user prompts from cxalloy to correctly identify different asset attributes using difflib. Each potential option for the attributes is refreshed each time from the cxalloy webpage automatically for the specific user selected project and utilized for attribute identification and labeling during the upload process. 


All three scripts will need a username and password input when the script is run with the username as the first field and password as the second. 

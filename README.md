# Automate PastPerfect

## Overview

The "Automate_PastPerfect" repository contains a project designed to automate the process of data entry into the Marist College Archives and Special Collections database, specifically within the PastPerfect application. This automation tool focuses on efficiently importing large datasets from an Excel sheet, significantly streamlining the data entry process.

Initially, this task was performed manually, a process that could potentially span several months. With the development of this automation tool, the data entry process is not only accelerated but also allows staff members to focus on other important tasks.

## Key Features

- **Automated Data Entry**: The tool reads data from an Excel file and automatically inputs it into the PastPerfect database.
- **Selenium Integration**: Utilizes Selenium for browser automation to interact with the PastPerfect web application.
- **Openpyxl for Excel Handling**: Employs openpyxl, a Python library, for reading from and writing to Excel files.

## Main Technologies

- **Openpyxl**: A Python library used for reading from and writing to Excel files.
- **Selenium**: An open-source umbrella project that provides tools and libraries for browser automation.

## Core Functionality

The automation is achieved through a Python script that utilizes Selenium and Openpyxl. The script performs the following tasks:

1. **Login to PastPerfect**: Automates the login process using credentials.
2. **Navigate to Relevant Sections**: Finds and clicks on the necessary buttons and links to navigate the PastPerfect interface.
3. **Read Data from Excel**: Opens and reads data from a specified Excel file.
4. **Data Entry**: Automatically inputs data into the correct fields in the PastPerfect application.
5. **Image Management**: Handles the uploading and management of associated images.

## Example Code

Here is a snippet from the main script used in this project:

```python
# Import necessary libraries
from openpyxl import load_workbook
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
# (additional imports...)

# Setup and configurations
PATH = "C:\AutomationTool\chromedriver"
driver = webdriver.Chrome(PATH)
# (additional setup code...)

# Main script to automate data entry
# (detailed script that includes login, data reading, data entry automation...)

# End of script
print("Task reached end")
```
This code represents the core functionality of the automation tool, showcasing how it interacts with the PastPerfect application and

manages the data entry process.

Installation and Usage
To use this tool, follow these steps:

Install Required Libraries: Make sure to have Selenium, webdriver-manager, and openpyxl installed. You can install them using pip:

bash
Copy code
pip install selenium webdriver-manager openpyxl
Configure Paths: Adjust the ChromeDriver and Excel file paths in the script to match your local setup.

Run the Script: Execute the script to start the automation process. Ensure that the PastPerfect application is accessible and your machine meets the necessary requirements for running Selenium scripts.

Contributing
Contributions to the project are welcome. If you have suggestions for improvements or encounter any issues, please feel free to open an issue or submit a pull request.

Acknowledgements
This project was developed to assist the staff at Marist College Archives and Special Collections, with the goal of improving efficiency and productivity.

Author
Connor Johnson - Initial work and development

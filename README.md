
## Accessibility Checker
This is a Python command-line tool for checking the accessibility of web pages. 

The tool uses the axe_selenium_python library to perform automated accessibility testing on a list of web pages provided in an Excel file. 

The tool generates a detailed accessibility report for each page and saves the results in a ZIP file along with an Excel file summarizing the results.


## Installation

Before running the tool, you will need to install the following dependencies:

- Python 3.x
- openpyxl library (pip install openpyxl)
- axe_selenium_python library (pip install axe_selenium_python)
- A web driver for your chosen browser (e.g., Firefox, Chrome). You can download the appropriate driver for your system from the following links:
  - Firefox: https://github.com/mozilla/geckodriver/releases
  - Chrome: https://sites.google.com/a/chromium.org/chromedriver/downloads
  - After downloading the web driver, make sure that the driver executable is in your system PATH.

  You can use requirements.txt from this repositories


## Usage
To use the tool, follow these steps:


1. Set the required environment variables:
  - EMAIL_ADDRESS: your email address for sending the report (optional)
- EMAIL_PASSWORD: your email password (optional)
- SMTP_SERVER: the SMTP server for your email provider (optional)
- SMTP_PORT: the SMTP port for your email provider (optional)
2. Edit an Excel file (name: pages.xlsx) with a list of web page URLs that you want to check. The URLs should be listed in the first column of the first sheet in the Excel file.
3. Run the tool by executing the run.py script with Python: ```python run.py```.
4. The tool will then start checking each web page in the Excel file and generate a detailed accessibility report for each page. The reports will be saved in a subdirectory called "pages" and a summary of the results will be saved in an Excel file called "results.xlsx".
5. If you choose to send the report by email, the tool will prompt you to enter the email address where you want to receive the report. The report will be sent as a ZIP file attachment.

## Acknowledgements

This tool was developed using the axe_selenium_python library created by Mozilla Services. The axe_selenium_python library provides an implementation of the Axe accessibility engine for Selenium-based testing.

You can find more information about the axe_selenium_python library and its source code on the following GitHub repository:

 - [Axe Selenium](https://github.com/mozilla-services/axe-selenium-python)



# webscraping related imports
from selenium import webdriver
from selenium.common import exceptions
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait, Select
from webdriver_manager.chrome import ChromeDriverManager
import time

# data management/manipulation related imports
import pandas as pd
from datetime import datetime, date
import math
import re
import os

# other imports
import subprocess


## setting the working directory to be the folder this file is located in
# Get the absolute path of the current Python file
file_path = os.path.abspath(__file__)

# Get the directory of the current Python file
file_directory = os.path.dirname(file_path)

# set the working directory to be the directory of this file
os.chdir(file_directory)

print(os.getcwd())

## File path definitions
test_auto_data_xlsx = f"{file_directory}\\test_auto_data1.xlsx" # defines the path to the input dataset



"""
-------------------------
Useful functions
"""

def export_auto_dataset(output_df, num_datalines_to_export):
    auto_dataset_for_export = output_df.head(num_datalines_to_export) # get the given number of lines from the start
    auto_dataset_for_export.to_csv("scraped_auto_premium.csv", index=False)

"""
-------------------------
"""


def auto_scape_all():
    # define a list of insurance companies to iterate through
    insurance_companies = ["AA", "AMI", "Tower"]

    # read in the test dataset
    test_auto_data = pd.read_excel(test_auto_data_xlsx, dtype={"Postcode":"int"})

    # hardcoding num cars
    num_cars = len(test_auto_data)

    # estimate the number of seconds testing all cars on each company website will take
    approximate_total_times = [(time * 2000) for time in [50, 40, 65]]
    total_time_hours = max(approximate_total_times) / 3600 # convert seconds to hours for the estimated longest time
    total_time_minutes = round((total_time_hours - int(total_time_hours)) * 60) # reformat into minute and hours
    total_time_hours = math.floor(total_time_hours)

    print(f"Program will take approximately {total_time_hours} hours and {total_time_minutes} minutes to scrape the premiums for {num_cars} cars on the AA, AMI and Tower Websites", end="\n\n\n")

    # Start all the processes
    processes = [subprocess.Popen(['python', f'Individual-company_scraper-files\\insurance-premium_web-scraping_{company}.py']) for company in insurance_companies]

    # Wait for all the processes to complete
    for process in processes:
        process.wait()


    # read the individual company datasets into the dictionary
    insurance_company_dfs = {f"{company}_df":pd.read_csv(f'Individual-company_data-files\\{company.lower()}_scraped_auto_premiums.csv') for company in insurance_companies}

    # temporary dataframe length adjustment
    test_auto_data = test_auto_data.head(2)

    # handling cases of adjusted agreed value

    # merging the individual datasets into one dataset for output
    for df in insurance_company_dfs.values():
        test_auto_data = pd.merge(test_auto_data.head(2), df, on="Sample Number")
    
    export_auto_dataset(test_auto_data, len(test_auto_data))


def main():

    # scrape all of the insurance premiums for the given car-person combinations
    auto_scape_all()


main()
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
import subprocess as sp
import multiprocessing

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
# define function to export the final dataset
def export_auto_dataset(output_df, num_cars):
    auto_dataset_for_export = output_df.head(num_cars) # get the given number of lines from the start
    auto_dataset_for_export.to_csv("scraped_auto_premium.csv", index=False)

# Define a function to run the subprocess
def run_subprocess(args):
    # saving the arguments as more intuitive names
    company, num_cars = args[0], args[1]

    # running the process (the python file)
    process = sp.Popen(['python', f'Individual-company_scraper-files\\insurance-premium_web-scraping_{company}.py'], stdin=sp.PIPE, text=True)
    process.communicate(input=str(num_cars)) # passing num_cars into the process as a standard input

def merge_dfs(dfs_list, merge_key_columns):
    # merging the individual datasets into one dataset for output
    merged_df = pd.merge(dfs_list[0], dfs_list[1], on=merge_key_columns)

    # if there are still more dfs to merge
    if len(dfs_list) >= 3:
        for df in dfs_list[2:]:
            merged_df = pd.merge(merged_df, df, on=merge_key_columns)

    return merged_df

def changes_in_agreed_value_adjustments(adjusted_df, insurance_companies):
    # iterates through all of the examples where the adjusted value had to be changed and ensures that the rest of the companies are using the same adjusted value
    for index, row in adjusted_df.iterrows():
        adjusted_down_values = []
        adjusted_up_values = []
        for company in insurance_companies:

            if row[f"{company}_agreed_value_was_adjusted"] == 1: # if the agreed value was adjusted UPWARDS for this company on this row
                print(f"agreed value was adjusted UPWARDS for {index}th example of {company}")
                print(row)
                adjusted_up_values.append(row[f"{company}_agreed_value"]) # saves the adjusted agreed value for later comparison
            elif row[f"{company}_agreed_value_was_adjusted"] == -1: # if the agreed value was adjusted DOWNWARDS for this company on this row
                print(f"agreed value was adjusted DOWNWARDS for {index}th example of {company}")
                adjusted_down_values.append(row[f"{company}_agreed_value"]) # saves the adjusted agreed value for later comparison
                print(row)

        
        if len(adjusted_up_values) > 0 and len(adjusted_down_values) > 0:
            print("INCONSISTENT AGREED VALUE RANGES", row)
        elif len(adjusted_up_values) > 0: # if agreed values were only adjusted up, select the smallest one (as it will be inside the accepted ranges of all of the others)
            print("Selected:", min(adjusted_up_values))
        elif len(adjusted_down_values) > 0: # if agreed values were only adjusted down, select the largest one (as it will be inside the accepted ranges of all of the others)
            print("Selected:", max(adjusted_down_values))
"""
-------------------------
"""


def auto_scape_all():
    # define a list of insurance companies to iterate through
    insurance_companies = ["AA", "AMI", "Tower"]

    # read in the test dataset
    test_auto_data = pd.read_excel(test_auto_data_xlsx, dtype={"Postcode":"int"})

    # saveing the length of test_auto_data as num_cars
    num_cars = 10
    #num_cars = len(test_auto_data)

    # estimate the number of seconds testing all cars on each company website will take
    approximate_total_times = [(time * num_cars) for time in [46, 40, 65]]
    total_time_hours = max(approximate_total_times) / 3600 # convert seconds to hours for the estimated longest time
    total_time_minutes = round((total_time_hours - int(total_time_hours)) * 60) # reformat into minute and hours
    total_time_hours = math.floor(total_time_hours)

    # print out the time to execute estimate
    print(f"Program will take approximately {total_time_hours} hours and {total_time_minutes} minutes to scrape the premiums for {num_cars} cars on the AA, AMI and Tower Websites", end="\n\n\n")

    
    #args = [(company, num_cars) for company in insurance_companies]

    # Start all the processes
    #with multiprocessing.Pool() as pool:
    #    pool.map(run_subprocess, args)


    # read the individual company datasets into the dictionary
    insurance_company_dfs = {f"{company}":pd.read_csv(f'Individual-company_data-files\\{company.lower()}_scraped_auto_premiums.csv') for company in insurance_companies}

    # merging the company datasets into a single dataset
    merged_company_df = merge_dfs(list(insurance_company_dfs.values()), ["Sample Number", "PolicyStartDate"])

    # gets all the examples where for any one of the companies, the agreed value was adjusted
    filtered_merged_company_df = merged_company_df.query("AA_agreed_value_was_adjusted != 0 or AMI_agreed_value_was_adjusted != 0 or Tower_agreed_value_was_adjusted != 0")

    # ensures that on all website, the agreed value input has been the same
    changes_in_agreed_value_adjustments(filtered_merged_company_df , insurance_companies)

    # temporary dataframe length adjustment
    #test_auto_data = test_auto_data.head(num_cars)

    #export_auto_dataset(test_auto_data, num_cars)


def main():

    # scrape all of the insurance premiums for the given car-person combinations
    auto_scape_all()

# run main() and ensure that it is only run when the code is called directly
if __name__ == "__main__":
    main()
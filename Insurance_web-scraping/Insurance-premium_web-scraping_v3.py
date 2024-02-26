## library imports
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

## File path definitions
test_auto_data_xlsx = f"{file_directory}\\test_auto_data1.xlsx" # defines the path to the input dataset



"""
-------------------------
Useful functions
"""
# define a function to merge dfs
def merge_dfs(dfs_list, merge_key_columns):
    # merging the individual datasets into one dataset for output
    merged_df = pd.merge(dfs_list[0], dfs_list[1], on=merge_key_columns)

    # if there are still more dfs to merge
    if len(dfs_list) >= 3:
        for df in dfs_list[2:]:
            merged_df = pd.merge(merged_df, df, on=merge_key_columns)

    return merged_df

# define function to export the final dataset
def export_auto_dataset(num_cars):
    # read in all of the individual company datasets
    individual_company_dfs = {company:pd.read_csv(f'Individual-company_data-files\\{company.lower()}_scraped_auto_premiums.csv') for company in insurance_companies}

    # reduce the length of test_auto_data to be equal to the specified number of cars
    output_df = test_auto_data.head(num_cars)

    # set the policy start date to be equal to be todays date (as all of the individual company dfs have the polciy start date as todays date)
    output_df.loc[:, 'PolicyStartDate'] = individual_company_dfs['AA']['PolicyStartDate'].values[0]

    # removing columns from the individual company dfs that we dont want in the final output
    individual_company_dfs = [individual_company_dfs[company].drop(["PolicyStartDate", f"{company}_agreed_value"], axis=1) for company in insurance_companies]

    # adding the info that was scraped from the website into the output dataframe
    output_df = merge_dfs([output_df] + individual_company_dfs, ["Sample Number"])

    # export the dataframe to the csv
    output_df.to_csv("scraped_auto_premiums.csv", index=False) 

# defining a function to run a single subprocess, is called by 'running_the_subprocesses'
def run_subprocess(args):
    # saving the arguments as more intuitive names
    company, row_indexes = args[0], args[1]
    # running the process (the python file)
    process = sp.Popen(['python', f'Individual-company_scraper-files\\insurance-premium_web-scraping_{company}.py'], stdin=sp.PIPE, text=True)

    process.communicate(input=str(row_indexes)) # passing all of the row indexes into the process as a standard input

# defining a function to run all the subprocesses
def runnning_the_subprocesses(indexes):
    # defining the arugments to be passed down into the subprocesses
    args = []
    for company in insurance_companies:
        if len(indexes[company]) > 0:
            # if there are at least 2 row to scrape for tower, split tower into 2 seperate processes
            if company == "Tower" and len(indexes["Tower"]) >= 2:
                midpoint = (len(indexes["Tower"]) + 1) // 2
                args.append( ("Tower", indexes["Tower"][:midpoint]) )
                args.append( ("Tower", indexes["Tower"][midpoint:]) )
            else:
                args.append((company, indexes[company]))

    # Start all the processes
    with multiprocessing.Pool() as pool:
        pool.map(run_subprocess, args)

# define a function to go back over all the examples where the agreed value was changed, as redo the scraping for the other companies, to ensure all have the same agreed value
def redo_changed_agreed_value(start_i, end_i):
    # checks if the agreed values were changed on any of the companies. If they were returns the row number
    def check_agreed_values(index):
        # saves the original agreed value (from the input excel spreadsheet) as a variable for comparison
        original_agreed_val = int(test_auto_data.loc[index, "AgreedValue"])

        # save all of the agreed values that are different from the original agreed value
        adjusted_agreed_val_companies = [f"{company}" for company in insurance_companies if original_agreed_val != int(insurance_company_dfs[f"{company}"].loc[index, f"{company}_agreed_value"])]

        if len(adjusted_agreed_val_companies) > 0:
            return index, adjusted_agreed_val_companies

    def selected_agreed_value():
        # saves the original agreed value (from the input excel spreadsheet) as a variable for comparison
        original_agreed_val = int(test_auto_data.loc[row_index, "AgreedValue"])

        if abs(original_agreed_val - agreed_value_maximum) < abs(original_agreed_val - agreed_value_minimum):
            return agreed_value_maximum
        else:
            return agreed_value_minimum

    # read all of the individual company datasets into the dictionary
    insurance_company_dfs = {f"{company}":pd.read_csv(f'Individual-company_data-files\\{company.lower()}_scraped_auto_premiums.csv') for company in insurance_companies}


    # gets all the row indexes where the agreed value has been modified
    adjusted_row_info = [(check_agreed_values(i)) for i in range(start_i, end_i) if check_agreed_values(i) != None]

    redo_scrape_args = {f"{company}":[] for company in insurance_companies}

    for row_index, companies in adjusted_row_info:
        # gets all the limits of the agreed values for each individual company from their dataframe
        agreed_value_maximum = [x for x in [int(insurance_company_dfs[f"{company}"].loc[row_index, f"{company}_agreed_value_maximum"]) for company in insurance_companies] if x >= 0] # gets all (that are non-negative) upper limits for the agreed value (as negative values can only occur if there was an error)
        agreed_value_minimum = [y for y in [int(insurance_company_dfs[f"{company}"].loc[row_index, f"{company}_agreed_value_minimum"]) for company in insurance_companies] if y >= 0] # gets all (that are non-negative) lower limits for the agreed value (as negative values can only occur if there was an error)

        # the upper and lower limits for agreed values are the lowest maximum and highest minimum
        agreed_value_maximum = min(agreed_value_maximum)
        agreed_value_minimum = max(agreed_value_minimum)

        # checking if the agreed value limits are not viable
        if agreed_value_maximum < agreed_value_minimum: # if the max is smaller than the min, then there are no agreed values that can be acceptable
            raise ValueError(f"Agreed value acceptable range is inconsistent on row {row_index}")
        
        # save the agreed values to test_auto_data
        chosen_agreed_value = selected_agreed_value()
        test_auto_data.loc[row_index, "AgreedValue"] = chosen_agreed_value

        for company in insurance_companies:
            if int(insurance_company_dfs[f"{company}"].loc[row_index, f"{company}_agreed_value"]) != chosen_agreed_value: 
                redo_scrape_args[company].append(row_index)
    
    # output a csv which the subprocesses will read from, including all of the newly modified agreed values
    test_auto_data.to_csv("test_auto_data1.csv", index=False) 
    
    # redo all of the scrapes where the agreed value was inconsistent
    runnning_the_subprocesses(redo_scrape_args)

# define a function to go and attempt to scrape from all the examples where there was an error that might be fixable
def redo_website_scrape_errors(start_i, end_i):
    # checks if the agreed values were changed on any of the companies. If they were returns the row number
    def find_fixable_errors(company):
        return [row_i for row_i in range(start_i, end_i) if insurance_company_dfs[f"{company}"].loc[row_i, f"{company}_Error_code"] == "Unknown Error"]

    # read all of the individual company datasets into the dictionary
    insurance_company_dfs = {f"{company}":pd.read_csv(f'Individual-company_data-files\\{company.lower()}_scraped_auto_premiums.csv') for company in insurance_companies}

    # gets all of the rows for each company that can be redone 
    redo_scrape_args = {f"{company}":find_fixable_errors(company) for company in insurance_companies}

    # redo all of the scrapes where the the error was 'unknown' meaning it could work on this attempt
    runnning_the_subprocesses(redo_scrape_args)

# deletes all the individual comapny csvs when we are finished with them (to make a blank slate for next time)
def delete_intermediary_csvs():
    file_csvs = [f"{file_directory}\\Individual-company_data-files\\{company.lower()}_scraped_auto_premiums.csv" for company in insurance_companies] + [f"{file_directory}\\test_auto_data1.csv"]
    for file_csv in file_csvs:
        os.remove(file_csv)

"""
-------------------------
"""


def auto_scape_all():
    # define a list of insurance companies to iterate through
    global insurance_companies
    insurance_companies = ["AA", "AMI", "Tower"]

    # read in the test dataset
    global test_auto_data
    test_auto_data = pd.read_excel(test_auto_data_xlsx, dtype={"Postcode":"int"})
    test_auto_data.to_csv("test_auto_data1.csv", index=False) # output a csv which the subprocesses will read from


    # saving the length of test_auto_data as num_cars
    num_cars = 50
    #num_cars = len(test_auto_data)

    # estimate the number of seconds testing all cars on each company website will take
    approximate_total_times = [(time * num_cars) for time in [46, 40, 70/2]]
    total_time_hours = max(approximate_total_times)*1.25 / 3600 # convert seconds to hours for the estimated longest time
    total_time_minutes = round((total_time_hours - int(total_time_hours)) * 60) # reformat into minute and hours
    total_time_hours = math.floor(total_time_hours)

    # print out the time to execute estimate
    print(f"Program will take approximately {total_time_hours} hours and {total_time_minutes} minutes to scrape the premiums for {num_cars} cars on the AA, AMI and Tower Websites", end="\n\n\n")

    # defining the row indexes of all of the cars we are going to scrape
    indexes = [i for i in range(num_cars)]
    indexes = {company:indexes for company in insurance_companies}

    # running all of the subprocesses (starting scraping from all website simultaneously)
    runnning_the_subprocesses(indexes)
    
    # for all errors that are potentially still scrapable, attempt to scrape again (to fix)
    redo_website_scrape_errors(0, num_cars)

    # ensures that on all website, the agreed value input has been the same
    redo_changed_agreed_value(0, num_cars)

    # export the dataset to 'scraped_auto_premiums'
    export_auto_dataset(num_cars)

    # delete intermediary (individual company) csvs
    delete_intermediary_csvs()

def main():

    # scrape all of the insurance premiums for the given car-person combinations
    auto_scape_all()

# run main() and ensure that it is only run when the code is called directly
if __name__ == "__main__":
    main()
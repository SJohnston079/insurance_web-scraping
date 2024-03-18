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
from dateutil.relativedelta import relativedelta
import math
import os
import sys

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
test_auto_data_csv = f"{file_directory}\\test_home_data.csv" # defines the path to the input dataset



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
def export_auto_dataset(indexes):
    # read in all of the individual company datasets
    individual_company_dfs = {company:pd.read_csv(f'Individual-company_data-files\\{company.lower()}_scraped_auto_premiums.csv') for company in insurance_companies}

    # reduce the length of test_auto_data to be equal to the specified number of cars
    output_df = test_auto_data.iloc[indexes]

    # setting output_df 'Sample Number' to be an integer
    output_df.loc[:,'Sample Number'] = output_df.loc[:,'Sample Number'].astype('int64')

    # converting all missing values into empty strings
    output_df = output_df.fillna('')

    # set the policy start date to be equal to be todays date (as all of the individual company dfs have the polciy start date as todays date)
    output_df.loc[:, 'PolicyStartDate'] = individual_company_dfs[insurance_companies[0]]['PolicyStartDate'].values[0]

    # removing columns from the individual company dfs that we dont want in the final output
    individual_company_dfs = [individual_company_dfs[company].drop(["PolicyStartDate", f"{company}_agreed_value"], axis=1) for company in insurance_companies]

    # adding the info that was scraped from the website into the output dataframe
    output_df = merge_dfs([output_df] + individual_company_dfs, ["Sample Number"])

    # export the dataframe to the csv
    output_df.to_csv("scraped_auto_premiums.csv", index=False) 

# defining the optimal allocation of the examples we need to scrape for each company, onto up to 10 chromedriver windows
def optimal_allocation(indexes_dict):
    # defining a variable to store the window allocations (calculates the approx time for each company to run, given each has only 1 window)
    approximate_total_times = {company:len(company_indexes)*company_times[company] for company, company_indexes in indexes_dict.items()}

    # initialising window allocation variable
    window_allocations = {company:1 for company in insurance_companies}

    # allocate the number of windows
    while sum(window_allocations.values()) < int(max_windows):
        # find the company which has the highest approximated time to execute
        max_time_company = max(approximate_total_times, key=approximate_total_times.get)

        # assign an extra window to the company which has the highest approximated time to execute
        window_allocations[max_time_company] += 1

        # recalculate the time to execute for the company which has had another window allocated to it
        approximate_total_times[max_time_company] *= (window_allocations[max_time_company]-1)/window_allocations[max_time_company]
    

    # ensuring that each company has enough examples to split into its assigned number of windows
    for company in insurance_companies:
        if window_allocations[company] > len(indexes_dict[company]): # if the number of windows assigned to the company is more than the number of examples that company has to do
            window_allocations[company] = len(indexes_dict[company])

    return window_allocations

# defining a function to run a single subprocess, is called by 'running_the_subprocesses'
def run_subprocess(args):
    # saving the arguments as more intuitive names
    use_registration_number, company, row_indexes = args[0], args[1], args[2]
    # running the process (the python file)
    process = sp.Popen(['python', f'Individual-company_scraper-files\\insurance-premium_web-scraping_{company}.py'], stdin=sp.PIPE, text=True)

    process.communicate(input=str( [use_registration_number]+row_indexes )) # passing all of the row indexes into the process as a standard input

# defining a function to run all the subprocesses
def runnning_the_subprocesses(indexes_dict):
    # getting the optimal allocation of companies to a number of windows to scrape the fastest
    optimal_window_allocation = optimal_allocation(indexes_dict)

    # defining the arugments to be passed down into the subprocesses
    args = []
    for company in insurance_companies:
        # define for each company, windows they should scrape from, by defining the indexes each are responsible for
        num_windows = optimal_window_allocation[company]
        num_examples = len(indexes_dict[company])

        for i in range(num_windows):
            start_index = (i*num_examples)//num_windows
            end_index = ((i+1)*num_examples)//num_windows

            if i == 0: # if is the 1st window
                args.append( (company, indexes_dict[company][:end_index]) )
            elif i + 1 == num_windows: # if is the last window
                args.append( (company, indexes_dict[company][start_index:]) )
            else: # if is a window in the middle
                args.append( (company, indexes_dict[company][start_index:end_index]) )
        '''
        if len(indexes_dict[company]) > 0:
            # if there are at least 2 row to scrape for tower, split tower into 2 seperate processes
            if company == "Tower" and len(indexes_dict["Tower"]) >= 2:
                midpoint = (len(indexes_dict["Tower"]) + 1) // 2
                args.append( (use_registration_number, "Tower", indexes_dict["Tower"][:midpoint]) )
                args.append( (use_registration_number, "Tower", indexes_dict["Tower"][midpoint:]) )
            else:
                args.append((use_registration_number, company, indexes_dict[company]))
        '''

    # Start all the processes
    with multiprocessing.Pool() as pool:
        pool.map(run_subprocess, args)

# define a function to go back over all the examples where the agreed value was changed, as redo the scraping for the other companies, to ensure all have the same agreed value
def redo_changed_agreed_value(start_i, end_i, indexes):
    # checks if the agreed values were changed on any of the companies. If they were returns the row number
    def check_agreed_values(i):
        # saves the original agreed value (from the input excel spreadsheet) as a variable for comparison
        original_agreed_val = int(test_auto_data.loc[indexes[i], "AgreedValue"])

        # save all of the agreed values that are different from the original agreed value
        adjusted_agreed_val_companies = [f"{company}" for company in insurance_companies if original_agreed_val != int(insurance_company_dfs[f"{company}"].loc[i, f"{company}_agreed_value"])]

        if len(adjusted_agreed_val_companies) > 0:
            return i, adjusted_agreed_val_companies

    def selected_agreed_value():
        # saves the original agreed value (from the input excel spreadsheet) as a variable for comparison
        original_agreed_val = int(test_auto_data.loc[indexes[row_index], "AgreedValue"])

        if abs(original_agreed_val - agreed_value_max) < abs(original_agreed_val - agreed_value_min):
            return agreed_value_max
        else:
            return agreed_value_min

    # read all of the individual company datasets into the dictionary
    insurance_company_dfs = {f"{company}":pd.read_csv(f'Individual-company_data-files\\{company.lower()}_scraped_auto_premiums.csv') for company in insurance_companies}


    # gets all the row indexes where the agreed value has been modified
    adjusted_row_info = [(check_agreed_values(i)) for i in range(start_i, end_i) if check_agreed_values(i) != None]

    redo_scrape_args = {f"{company}":[] for company in insurance_companies}

    for row_index, companies in adjusted_row_info:
        # gets all the limits of the agreed values for each individual company from their dataframe
        agreed_value_maximums = [x for x in [int(insurance_company_dfs[f"{company}"].loc[row_index, f"{company}_agreed_value_maximum"]) for company in insurance_companies] if x >= 0] # gets all (that are non-negative) upper limits for the agreed value (as negative values can only occur if there was an error)
        agreed_value_minimums = [y for y in [int(insurance_company_dfs[f"{company}"].loc[row_index, f"{company}_agreed_value_minimum"]) for company in insurance_companies] if y >= 0] # gets all (that are non-negative) lower limits for the agreed value (as negative values can only occur if there was an error)

        # the upper and lower limits for agreed values are the lowest maximum and highest minimum
        if len(agreed_value_maximums) > 0: 
            agreed_value_max = min(agreed_value_maximums)
            agreed_value_min = max(agreed_value_minimums)
        
            # checking if the agreed value limits are not viable
            if agreed_value_max < agreed_value_min: # if the max is smaller than the min, then there are no agreed values that can be acceptable
                print(f"Agreed value acceptable range is inconsistent on row {indexes[row_index]}")
                test_auto_data.loc[row_index, f"General_Error_Code"] = "Agreed value ranges are inconsistent"
            
            # save the agreed values to test_auto_data
            chosen_agreed_value = selected_agreed_value()
            test_auto_data.loc[indexes[row_index], "AgreedValue"] = chosen_agreed_value

            # for each company, save the rows for which their current agreed value is not consistent with the chosen agreed value. We then repeat the scraping for these examples with the new agreed value
            for company in insurance_companies:
                if int(insurance_company_dfs[f"{company}"].loc[row_index, f"{company}_agreed_value"]) != chosen_agreed_value and insurance_company_dfs[f"{company}"].loc[row_index, f"{company}_agreed_value_maximum"] != -1: # if agreed value is not consistent with the new agreed value AND the agreed value limits are not invalid
                    redo_scrape_args[company].append(indexes[row_index])
        
    # output a csv which the subprocesses will read from, including all of the newly modified agreed values
    test_auto_data.to_csv("test_auto_data1.csv", index=False) 
    
    # redo all of the scrapes where the agreed value was inconsistent
    runnning_the_subprocesses(redo_scrape_args)

# define a function to go and attempt to scrape from all the examples where there was an error that might be fixable
def redo_website_scrape_errors(start_i, end_i, indexes):
    # checks if the agreed values were changed on any of the companies. If they were returns the row number
    def find_fixable_errors(company):
        return [indexes[row_i] for row_i in range(start_i, end_i) if insurance_company_dfs[f"{company}"].loc[row_i, f"{company}_Error_code"] == "Unknown Error"]

    # read all of the individual company datasets into the dictionary
    insurance_company_dfs = {f"{company}":pd.read_csv(f'Individual-company_data-files\\{company.lower()}_scraped_auto_premiums.csv') for company in insurance_companies}

    # gets all of the rows for each company that can be redone 
    redo_scrape_args = {f"{company}":find_fixable_errors(company) for company in insurance_companies}

    # redo all of the scrapes where the the error was 'unknown' meaning it could work on this attempt
    runnning_the_subprocesses(redo_scrape_args)

# deletes all the individual comapny csvs when we are finished with them (to make a blank slate for next time)
def delete_intermediary_csvs():
    file_csvs = [f"{file_directory}\\Individual-company_data-files\\{company.lower()}_scraped_home_premiums.csv" for company in insurance_companies] + [f"{file_directory}\\test_home_data.csv"]
    for file_csv in file_csvs:
        try:
            os.remove(file_csv)
        except FileNotFoundError:
            pass

# performing dataset preprocessing that all companies will inherit
def dataset_preprocess():
    # read in the test dataset
    global test_auto_data
    test_auto_data = pd.read_csv(test_auto_data_csv, dtype={"Postcode":"int"})

    # converts all variables to type string
    test_auto_data = test_auto_data.astype(str)
    test_auto_data.fillna("") # setting all NA values to be an empty string

    # output a csv which the subprocesses will read from
    test_auto_data.to_csv("test_auto_data1.csv", index=False) 
"""
-------------------------
"""


def auto_scape_all():
    # customise what insurance companies to scrape from
    def select_insurance_companies():
        print("Enter 1 to include the website, 0 to exclude:", end="\n------------------------------------------------------------------\n")
        selected_insurance_companies = []
        for i in range(len(insurance_companies)):
            include_index = input(f"Include {insurance_companies[i]} in the websites to scrape from?: ")
            while include_index not in ["0","1"]:
                print("Try again, accepted values are 0:'not include', 1:'include'")
                include_index = input(f"Include {insurance_companies[i]} in the websites to scrape from?: ")

            if include_index == "1":
                selected_insurance_companies.append(insurance_companies[i])
        return selected_insurance_companies 

    # define a list of insurance companies to iterate through
    global insurance_companies
    insurance_companies = ["AA", "AMI", "Tower"]
    insurance_companies = select_insurance_companies()

    # checks to make sure at least 1 company was selected to scrape from
    if len(insurance_companies) == 0:
        raise Exception("Must Select at LEAST one company to scrape from")

    print(f"Scraping from {insurance_companies}", end="\n\n")

    # delete intermediary (individual company) csvs (ensuring blank slate start)
    #delete_intermediary_csvs()

    # read in the test dataset
    dataset_preprocess()


    ## getting a number for how many houses we wish to scrape from
    print("\nHow many houses do you wish scrape premiums for?")
    num_houses = ""

    # repeatedly prompt the user until is given a valid number of cars to scrape from
    while not isinstance(num_houses, int):
        num_houses = input(f"Enter EITHER (A number between 1 and {len(test_auto_data)} inclusive) OR ('ALL' to scrape for all examples in the test_home_data): ").upper()

        # attempting to convert the number of cars to an integer
        try: 
            if num_houses == "ALL": # check if num_houses is a valid string (set it to the number of examples in the dataset)
                num_houses = len(test_auto_data)
            else:
                num_houses = int(num_houses) # attempt to convert num_houses into an integer value
                
                # if num_houses is able to be converted into an integer, is the integer invalid (larger than the number of examples in the spreadsheet)
                if num_houses > len(test_auto_data):
                    raise ValueError
        except: # if the num_houses variable is invalid in some way, then set it to an empty string (to prompt again for input)
            num_houses = ""
    
    


    # defining the company times taken to scrape
    global company_times
    company_times = {"AA":50, "AMI":50, "Tower":90}
    company_times = {company:company_times[company] for company in insurance_companies} # redefing the variable to only save for companies that have been chosen
    

    # defining the row indexes of all of the cars we are going to scrape
    indexes = [i for i in range(num_houses)]
    indexes_dict = {company:indexes for company in insurance_companies}


    # defining the maximum number of allowed windows
    global max_windows
    max_windows = input(f"Choose the maximum number of browser windows to use (more can be quicker if your computers CPU can handle it (but not always)) (Recommended {len(insurance_companies)}-7):")
    while not max_windows.isdigit():
        max_windows = input("Must input only digits. The max number of browser windows to allow: ")

    # setting max_windows to be an integer
    max_windows = int(max_windows)

    # ensuring that there is at least 1 window per company that we wish to scrape from 
    if max_windows < len(insurance_companies):
        max_windows = len(insurance_companies)
    # printing newlines for more readable output
    print("\n\n")

    # calculating an estimate for the length of time the program needs to run
    window_allocations = optimal_allocation(indexes_dict)
    approximate_total_times = [time*(num_houses/window_allocations[company]) for company, time in company_times.items()]
    total_time_hours = max(approximate_total_times)*1.5 / 3600 # convert seconds to hours for the estimated longest time
    total_time_minutes = round((total_time_hours - int(total_time_hours)) * 60) # reformat into minute and hours
    total_time_hours = math.floor(total_time_hours)

    # constructing the string to print out our time estimate
    time_estimate_string = f"Program will take approximately {total_time_hours} hours and {total_time_minutes} minutes to scrape the premiums for {num_houses} cars on the "
    if len(insurance_companies) == 3:
        time_estimate_string += "AA, AMI and Tower Websites"
    if len(insurance_companies) == 2:
        time_estimate_string += f"{insurance_companies[0]} and {insurance_companies[1]} Websites"
    if len(insurance_companies) == 1:
        time_estimate_string += f"{insurance_companies[0]} Website"
    # print out the time to execute estimate
    print(time_estimate_string)

    # get start time
    start_time = time.time()

    # running all of the subprocesses (starting scraping from all website simultaneously)
    runnning_the_subprocesses(indexes_dict)
    
    end_time = time.time() # get time of end of each iteration
    print("Main subprocesses runtime:", round(end_time - start_time,2)) # print out the length of time taken

    # get start time
    start_time = time.time()

    # for all errors that are potentially still scrapable, attempt to scrape again (to fix)
    redo_website_scrape_errors(0, num_houses, indexes)

    end_time = time.time() # get time of end of each iteration
    print("Redoing scrape errors runtime:", round(end_time - start_time,2)) # print out the length of time taken

    # get start time
    start_time = time.time()

    # ensures that on all website, the agreed value input has been the same
    redo_changed_agreed_value(0, num_houses, indexes)

    end_time = time.time() # get time of end of each iteration
    print("Redoing inconsistent agreed values runtime:", round(end_time - start_time,2)) # print out the length of time taken

    # export the dataset to 'scraped_auto_premiums'
    export_auto_dataset(indexes)

    # delete intermediary (individual company) csvs (ensuring blank slate start)
    delete_intermediary_csvs()

def main():
    # scrape all of the insurance premiums for the given car-person combinations
    auto_scape_all()



# run main() and ensure that it is only run when the code is called directly
if __name__ == "__main__":
    main()
# webscraping related imports
from selenium import webdriver
from selenium.common import exceptions
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options

# data management/manipulation related imports
import pandas as pd
import re
import os
import time

"""
-------------------------
Useful functions
"""
def define_file_path(file_name = "test_home_data.xlsx"):
    ## setting the working directory to be the folder this file is located in
    # Get the absolute path of the current Python file
    file_path = os.path.abspath(__file__)

    # Get the directory of the current Python file
    file_directory = os.path.dirname(file_path)

    # Get the parent directory of the current Python file
    parent_dir = os.path.abspath(os.path.join(file_directory, os.pardir))

    ## File path definitions
    file_path = f"{parent_dir}\\{file_name}" # defines the path to the input dataset

    return file_path


def export_auto_dataset(output_df, test_home_data_df, company, input_indexes):
    # set the dates to be what was scraped
    output_df.loc[:,"PolicyStartDate"] = test_home_data_df.loc[:,"PolicyStartDate"]

    try:
        # set the column "Sample Number" to be the index column
        output_df.set_index("Sample Number", drop=True, inplace=True)

        # attempt to read in the output dataset
        insurance_premium_web_scraping_df = pd.read_csv(define_file_path(file_name = f"//Individual-company_data-files//{company}"))
    except FileNotFoundError:
        # if there exists no df already, then we just export what we have scraped
        auto_dataset_for_export = output_df.iloc[input_indexes] # get the given number of lines from the start
    else:
        # set the column "Sample Number" to be the index column so the dataframe combine on "Sample Number"
        insurance_premium_web_scraping_df.set_index("Sample Number", drop=True, inplace=True)

        # combine with the newly scraped data (anywhere there is an overlap, the newer (just scraped) data overwrites the older data)
        auto_dataset_for_export = output_df.iloc[input_indexes].combine_first(insurance_premium_web_scraping_df)

        # sort the dataset on the index
        auto_dataset_for_export.sort_index(inplace= True)
    finally:
        # export the dataset
        auto_dataset_for_export.to_csv(define_file_path(file_name = f"\\Individual-company_data-files\\{company}_scraped_auto_premiums.csv"))


def remove_non_numeric(string):
    return ''.join(char for char in string if (char.isdigit() or char == "."))


# defines a function to reformat the postcodes in test_auto_data
def postcode_reformat(postcode):
    postcode = str(postcode) # converts the postcode input into a string
    while len(postcode) != 4:
        postcode = f"0{postcode}"
    return postcode


def convert_money_str_to_int(money_string, cents = False):
    money_string = re.sub(r'\(.*?\)', '', money_string) # deletes all characters that are in between brackets (needed for tower as its annual payment includes '(save $88.37)' or other equivalent numbers)
    money_string = remove_non_numeric(money_string)
    if cents:
        return float(money_string)
    else:
        return int(money_string)


def load_webdriver():
    # setting up driver options settings
    chrome_options = Options()
    chrome_options.add_argument('--log-level=3')  # Set logging level to WARNING
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options) # loads chromedriver with the given options settings

    # define the implicit wait time for the session
    driver.implicitly_wait(1)

    # defines Wait and Wait10, dynamic wait templates
    Wait1 = WebDriverWait(driver, 1)
    Wait3 = WebDriverWait(driver, 3)
    Wait10 = WebDriverWait(driver, 10)

    return driver, Wait1, Wait3, Wait10


def read_indicies_to_scrape():
    ## reading in variables from the standard input that the 'parent' process passes in
    input_indexes = input()
    input_indexes = input_indexes.replace("[", "").replace("]", "").split(",")

    # return all of the row incidies to scrape from
    return list(map(int, input_indexes))


def choose_excess_value(excess_vals, desired_excess):
    # Initialize the closest value to a large number
    closest_value = float('inf')
    closest_excess = -1

    # Iterate over the list of excess values
    for i, excess_val in enumerate(excess_vals):
        if abs(excess_val - desired_excess) < abs(closest_value - desired_excess):
            closest_value = excess_val
            closest_index = i

    return excess_vals[closest_index]


# performing the data reading in and preprocessing for the given company
def dataset_preprocess(company):
    # read in the data
    test_home_data_df = pd.read_excel(define_file_path(), dtype={"Postcode":"int"})
    '''
    # setting all NA values in string columns to to be an empty string
    test_home_data_df = test_home_data_df.apply(lambda x: x.fillna("") if x.dtype == "object" else x)

    # setting some variables to be string types
    test_home_data_df.loc[:,"AgreedValue"] = test_home_data_df.loc[:,"AgreedValue"].astype(str)
    '''
    ## converts the two values which should be dates into dates
    # convert the date of birth variable into a date object
    test_home_data_df['DOB'] = pd.to_datetime(test_home_data_df['DOB'], format='%d/%m/%Y')
    '''
    # convert all the date of incident variables into a date objects
    for i in range(1,6):
        test_home_data_df[f'Date_of_incident{i}'] = pd.to_datetime(test_home_data_df[f'Date_of_incident{i}'], format='%Y/%m/%d')
 

    '''
    # pads out the front of postcodes with zeroes (as excel removes leading zeros)
    test_home_data_df['Postcode'] = test_home_data_df['Postcode'].apply(postcode_reformat)
    
    # define a pandas dataframe to output the results from this scraping
    output_df = test_home_data_df.loc[:, ["Sample Number", "Excess"]]
    '''
    output_df["AA_agreed_value"] = test_home_data_df["AgreedValue"].to_string(index=False).strip().split()
    output_df["AA_monthly_premium"] = ["-1"] * len(test_home_data_df)
    output_df["AA_yearly_premium"] = ["-1"] * len(test_home_data_df)
    output_df["AA_agreed_value_minimum"] = [-1] * len(test_home_data_df)
    output_df["AA_agreed_value_maximum"] = [-1] * len(test_home_data_df)
    output_df["AA_Error_code"] = ["No Error"] * len(test_home_data_df)
    output_df["AA_selected_car_variant"] = [""] * len(test_home_data_df)
    '''
    output_df[f"{company}_selected_address"] = [""] * len(test_home_data_df)
    return test_home_data_df, output_df

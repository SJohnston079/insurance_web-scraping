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
from datetime import datetime
from dateutil.relativedelta import relativedelta
import re
import os
import time

"""
-------------------------
Useful functions
"""
# function which returns the directory of the current file
def get_current_file_directory():
    # Get the absolute path of the current Python file
    file_path = os.path.abspath(__file__)

    # Get the directory of the current Python file
    file_directory = os.path.dirname(file_path)
    return file_directory

# defines the file path for the input dataset
def define_file_path():
    # get the directory of this current file
    file_directory = get_current_file_directory()

    # Get the parent directory of the current Python file
    parent_dir = os.path.abspath(os.path.join(file_directory, os.pardir))

    # defining the file path of the input dataset 'test_auto_data1.xlsx'
    file_path = f"{parent_dir}\\test_auto_data1.xlsx" # defines the path to the input dataset

    # returning the correct file path
    return file_path


# loading the webdriver
def load_webdriver():
    # loads chromedriver
    global driver # defines driver as a global variableaa
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))

    # setting up driver options settings
    chrome_options = Options()
    chrome_options.add_argument('--log-level=3')  # Set logging level to WARNING
    driver = webdriver.Chrome(options=chrome_options)

    # define the implicit wait time for the session
    driver.implicitly_wait(1)

    # defines Wait and Wait10, dynamic wait templates
    Wait3 = WebDriverWait(driver, 3)
    Wait10 = WebDriverWait(driver, 10)

    return driver, Wait3, Wait10


# exports the dataset
def export_auto_dataset(output_df, company):
    # get the directory of this current file
    file_directory = get_current_file_directory()

    # getting a string of todays date
    todays_date = datetime.today().strftime("%Y-%m-%d")

    # defining the name of the file we wish to output
    filename = f'{todays_date}_{company}_output_1'

    # ensuring that the new file doesn't overwrite the older file, by adding 1 to the number at the end until the file has a unqiue name
    index = 2
    while os.path.exists(f"{file_directory}\\Scraped_insurance-premium_outputs\\{filename}.csv"):
        filename = f"{filename[:-1]}{index}" # remove the previously used number, add a new one
        index += 1

        
    # output the scraped value to a csv
    output_df.to_csv(f"{file_directory}\\Scraped_insurance-premium_outputs\\{filename}.csv", index=False)


# takes in a string, and removes all of the non-numeric characters
def remove_non_numeric(string):
    return ''.join(char for char in string if (char.isdigit() or char == "."))


# defines a function to reformat the postcodes in test_auto_data
def postcode_reformat(postcode):
    postcode = str(postcode) # converts the postcode input into a string
    while len(postcode) != 4:
        postcode = f"0{postcode}"
    return postcode


# takes in a 'money string' e.g. "$4,000.56" and converts it to an integer (or float) e.g. 4000.56
def convert_money_str_to_int(money_string, cents = False): 
    money_string = re.sub(r'\(.*?\)', '', money_string) # deletes all characters that are in between brackets (needed for tower as its annual payment includes '(save $88.37)' or other equivalent numbers)
    modifided_money_string = remove_non_numeric(money_string)
    if modifided_money_string == "":
        raise Exception(f"Unknown Error: Scraped Premium '{money_string}' Wasn't Able To Be Converted To A Float/Int")
    elif cents:
        return float(modifided_money_string)
    else:
        return int(modifided_money_string)


# performing the data reading in and preprocessing
def dataset_preprocess(company):
    def incident_dates_calc(row):

        # defining a variable to store the output in
        incident_date_list = []

        # defining the number of incidient in the last 3 and 5 years for this person as a local variable
        num_incidents_3_year = row["Incidents_last3years_AA"]
        num_incidents_5_year = row["Incidents_last5years_AMISTATE"]


        # handle first incident date
        if num_incidents_5_year > 0:
            incident_date_list.append(datetime.strftime(row["PolicyStartDate"] - relativedelta(months=row["Month Since Last Claim"]), "%Y/%m/%d"))

            if num_incidents_3_year > 0:
                # calculate the dates for all incidents within 3 years
                step3year = (3*12-row["Month Since Last Claim"])//(num_incidents_3_year)
                for i in range(1, num_incidents_3_year):
                    incident_date_list.append(datetime.strftime(row["PolicyStartDate"] - relativedelta(months= step3year*i + num_incidents_3_year) - relativedelta(weeks=1), "%Y/%m/%d"))

                # calculate the dates for all incidents within 5 years
                num_incidents3_to_5_years = num_incidents_5_year - num_incidents_3_year

                if num_incidents3_to_5_years > 0:
                    step5year = ((5-3)*12)//(num_incidents3_to_5_years)
                    for i in range(1, num_incidents3_to_5_years + 1):
                        incident_date_list.append(datetime.strftime(row["PolicyStartDate"] - relativedelta(months= step5year*i + 3*12) - relativedelta(weeks=1), "%Y/%m/%d"))
            else:
                # calculate the dates for all incidents within 5 years
                step5year = (5*12)//(num_incidents_5_year)
                for i in range(1, num_incidents_5_year + 1):
                    incident_date_list.append(datetime.strftime(row["PolicyStartDate"] - relativedelta(months= step5year*i) - relativedelta(weeks=1), "%Y/%m/%d"))



        # for all remaining incident date variables, set to empty string
        while len(incident_date_list) < 5:
            incident_date_list.append(pd.NaT)

        # returning the output
        return incident_date_list

    # defining file paths
    input_data_file_path = define_file_path()

    
    # read in the test dataset
    global test_auto_data_df
    test_auto_data_df = pd.read_excel(input_data_file_path, dtype={"Street_number":str, "Postcode":str, "DOB":"datetime64[ns]"})
    

    ## setting up the incident dates (given the number of incidents in the last 3 and 5 years, assigns each incident a date)
    incident_dates = test_auto_data_df.apply(lambda row: incident_dates_calc(row), axis=1)
    incident_dates = [list(x) for x in zip(*incident_dates)] # reformatting the output to allow the values to be saved to the dataframe

    for i in range(5):
        test_auto_data_df.loc[:,f"Date_of_incident{i+1}"] = pd.to_datetime(incident_dates[i], format="%Y/%m/%d")


    # setting all NA values in string columns to to be an empty string
    test_auto_data_df = test_auto_data_df.apply(lambda x: x.fillna("") if x.dtype == "object" else x)


    # creates a new dataframe to save the scraped info
    output_df = test_auto_data_df.loc[:, ["Sample Number", "PolicyStartDate"]]
    output_df[f"{company}_monthly_premium"] = ["-1"] * len(test_auto_data_df)
    output_df[f"{company}_yearly_premium"] = ["-1"] * len(test_auto_data_df)
    output_df[f"{company}_agreed_value_minimum"] = [-1] * len(test_auto_data_df)
    output_df[f"{company}_agreed_value_maximum"] = [-1] * len(test_auto_data_df)
    output_df[f"{company}_Error_code"] = ["No Error"] * len(test_auto_data_df)
    output_df[f"{company}_selected_car_variant"] = [""] * len(test_auto_data_df)
    output_df[f"{company}_selected_address"] = [""] * len(test_auto_data_df)
    output_df.set_index("Sample Number")

    # outputs the datasets
    return test_auto_data_df, output_df

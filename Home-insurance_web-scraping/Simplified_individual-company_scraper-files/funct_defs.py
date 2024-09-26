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
from datetime import datetime, timedelta
import re
import os

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
    file_path = f"{parent_dir}\\test_home_data.xlsx" # defines the path to the input dataset

    # returning the correct file path
    return file_path



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



# function which checks if the given date date_x occured within the last 'years' many years
def check_date_range(date_x, years):
    # Calculate the date 'years' many years ago
    y_years_ago = datetime.now() - timedelta(days=365*years)

    # Check if the input date is within the last 'years' many years
    if date_x >= y_years_ago:
        return True
    else:
        return False


# loading the webdriver
def load_webdriver():
    # loads chromedriver
    global driver # defines driver as a global variableaa
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))

    # setting up driver options settings
    chrome_options = Options()
    chrome_options.add_argument('--log-level=3', )  # Set logging level to WARNING
    chrome_options.add_argument("--disable-cookies")
    driver = webdriver.Chrome(options=chrome_options)

    # define the implicit wait time for the session
    driver.implicitly_wait(1)

    # defines Wait and Wait10, dynamic wait templates
    Wait1 = WebDriverWait(driver, 1)
    Wait3 = WebDriverWait(driver, 3)
    Wait10 = WebDriverWait(driver, 10)

    return driver, Wait1, Wait3, Wait10


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
    money_string = remove_non_numeric(money_string)
    if cents:
        return float(money_string)
    else:
        return int(money_string.replace(".", ""))


# define a function to help with preprocessing the address
def edit_street_names(row):
    if ' (' in row['Street_name']:
        street_name, street_type_modifier = row['Street_name'].split(' (')
        row['Street_name'] = street_name
        row['Street_type'] = f'{row["Street_type"]} {street_type_modifier.strip(")")}'
    return row



# performing the data reading in and preprocessing
def dataset_preprocess(company):
    # defining file paths
    input_data_file_path = define_file_path()

    # read in the test dataset
    global test_auto_data_df
    test_auto_data_df = pd.read_excel(input_data_file_path, dtype={"PolicyStartDate":'datetime64[ns]', "Street_name":str, "Postcode":int, "Occupancy":str, "ShortTermTenancy":str, 
                                                                   "DOB":'datetime64[ns]', "ConstructionType":str, "RoofType":str, "YearBuilt":str, "DwellingFloorArea":str, "GarageFloorArea":str, 
                                                                   "SumInsured":int, "Date_of_incident":'datetime64[ns]', "HouseSecurity":str, "MortgageeSale":str})
    
    # pads out the front of postcodes with zeroes (as excel removes leading zeros)
    test_auto_data_df['Postcode'] = test_auto_data_df['Postcode'].apply(postcode_reformat) 

    # removes all trailing and leading whitespace from the suburb variable (this may need to be done for other variables as well)
    test_auto_data_df.loc[:,'Suburb'] = test_auto_data_df.loc[:,'Suburb'].apply(lambda x: x.strip())

    # handles the street direction/ location. It is found within brackets e.g. King (West)
    test_auto_data_df = test_auto_data_df.apply(edit_street_names, axis=1)
    
    # creates a new dataframe to save the scraped info
    output_df = test_auto_data_df.loc[:, ["Sample Number", "PolicyStartDate", "Excess"]]
    output_df[f"{company}_monthly_premium"] = ["-1"] * len(test_auto_data_df)
    output_df[f"{company}_yearly_premium"] = ["-1"] * len(test_auto_data_df)
    output_df[f"{company}_Error_code"] = ["No Error"] * len(test_auto_data_df)
    output_df[f"{company}_selected_address"] = [""] * len(test_auto_data_df)
    
    # adding the variable 'Estimated replacement cost' only to the companies where it is relevant
    if company in ("AA"):
        output_df[f"{company}_estimated_replacement_cost"] = [-1] * len(test_auto_data_df)

    output_df.set_index("Sample Number")

    return test_auto_data_df, output_df

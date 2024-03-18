# webscraping related imports
from selenium import webdriver
from selenium.common import exceptions
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait, Select
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
import time

# data management/manipulation related imports
import pandas as pd
from datetime import datetime, date
import math
import re
import os
import sys


## setting the working directory to be the folder this file is located in
# Get the absolute path of the current Python file
file_path = os.path.abspath(__file__)

# Get the directory of the current Python file
file_directory = os.path.dirname(file_path)

# Get the parent directory of the current Python file
parent_dir = os.path.abspath(os.path.join(file_directory, os.pardir))

## File path definitions
test_auto_data_csv = f"{parent_dir}\\test_auto_data1.csv" # defines the path to the input dataset

"""
-------------------------
Useful functions
"""
def export_auto_dataset(row_indexes):

    # set the dates to be what was scraped
    global aa_output_df # making it so that this 'reset' of the index is done to the globally stored data
    aa_output_df.loc[:,"PolicyStartDate"] = test_auto_data_df.loc[:,"PolicyStartDate"]

    try:
        # set the column "Sample Number" to be the index column
        aa_output_df.set_index("Sample Number", drop=True, inplace=True)

        # read in the output dataset
        insurance_premium_web_scraping_AA_df = pd.read_csv(f"{parent_dir}\\Individual-company_data-files\\aa_scraped_auto_premiums.csv")
    
    except FileNotFoundError:
        # if there exists no df already, then we just export what we have scraped
        auto_dataset_for_export = aa_output_df.iloc[row_indexes] # get the given number of lines from the start
    else:
        # set the column "Sample Number" to be the index column so the dataframe combine on "Sample Number"
        insurance_premium_web_scraping_AA_df.set_index("Sample Number", drop=True, inplace=True)

        # combine with the newly scraped data (anywhere there is an overlap, the newer (just scraped) data overwrites the older data)
        auto_dataset_for_export = aa_output_df.iloc[row_indexes].combine_first(insurance_premium_web_scraping_AA_df)

        # sort the dataset on the index
        auto_dataset_for_export.sort_index(inplace= True)
    finally:
        # export the dataset
        auto_dataset_for_export.to_csv(f"{parent_dir}\\Individual-company_data-files\\aa_scraped_auto_premiums.csv")


def remove_non_numeric(string):
    return ''.join(char for char in string if (char.isdigit() or char == "."))


# defines a function to reformat the postcodes in test_auto_data
def postcode_reformat(postcode):
    postcode = str(postcode) # converts the postcode input into a string
    while len(postcode) != 4:
        postcode = f"0{postcode}"
    return postcode

    # reads in the test data for car insurance inputs


def convert_money_str_to_int(money_string, cents = False):
    money_string = re.sub(r'\(.*?\)', '', money_string) # deletes all characters that are in between brackets (needed for tower as its annual payment includes '(save $88.37)' or other equivalent numbers)
    money_string = remove_non_numeric(money_string)
    if cents:
        return float(money_string)
    else:
        return int(money_string)

# performing the data reading in and preprocessing
def dataset_preprocess():
    # read in the data
    global test_auto_data_df
    test_auto_data_df = pd.read_csv(test_auto_data_csv, dtype={"Postcode":"int"})

    # setting all NA values in string columns to to be an empty string
    test_auto_data_df = test_auto_data_df.apply(lambda x: x.fillna("") if x.dtype == "object" else x)

    # setting some variables to be string types
    test_auto_data_df.loc[:,"AgreedValue"] = test_auto_data_df.loc[:,"AgreedValue"].astype(str)

    ## converts the two values which should be dates into dates
    # convert the date of birth variable into a date object
    test_auto_data_df['DOB'] = pd.to_datetime(test_auto_data_df['DOB'], format='%Y-%m-%d')

    # convert all the date of incident variables into a date objects
    for i in range(1,6):
        test_auto_data_df[f'Date_of_incident{i}'] = pd.to_datetime(test_auto_data_df[f'Date_of_incident{i}'], format='%Y/%m/%d')

    # pads out the front of postcodes with zeroes (as excel removes leading zeros)
    test_auto_data_df['Postcode'] = test_auto_data_df['Postcode'].apply(postcode_reformat) 

    # creates a new dataframe to save the scraped info
    global aa_output_df
    aa_output_df = test_auto_data_df.loc[:, ["Sample Number", "PolicyStartDate"]]
    aa_output_df["AA_agreed_value"] = test_auto_data_df["AgreedValue"].to_string(index=False).strip().split()
    aa_output_df["AA_monthly_premium"] = ["-1"] * len(test_auto_data_df)
    aa_output_df["AA_yearly_premium"] = ["-1"] * len(test_auto_data_df)
    aa_output_df["AA_agreed_value_minimum"] = [-1] * len(test_auto_data_df)
    aa_output_df["AA_agreed_value_maximum"] = [-1] * len(test_auto_data_df)
    aa_output_df["AA_Error_code"] = ["No Error"] * len(test_auto_data_df)
    aa_output_df["AA_selected_car_variant"] = [""] * len(test_auto_data_df)
    aa_output_df["AA_selected_address"] = [""] * len(test_auto_data_df)


# a function to open the webdriver (chrome simulation)
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
    global Wait
    Wait = WebDriverWait(driver, 3)

    global Wait10
    Wait10 = WebDriverWait(driver, 10)

"""
-------------------------
"""

# defining a function that will scrape the premium for a single car from aa
def aa_auto_scrape(person_i):
    # defining a function which take the information from the spreadsheet and formats it so it can be used to scrape premium from aa website
    def aa_auto_data_format(person_i):

        # formatting model type
        model_type = test_auto_data_df.loc[person_i,'Type']

        if "XL Dual Cab" in model_type:
            model_type = model_type.replace("Dual", "Double").upper()
        else:
            model_type = model_type.upper()

        # formatting model series as a string
        model_series = test_auto_data_df.loc[person_i,'Series']

        # getting the street address
        house_number = remove_non_numeric( test_auto_data_df.loc[person_i,'Street_number'] ) # removes all non-numeric characters from the house number (e.g. removes A from 14A)
        street_name = test_auto_data_df.loc[person_i,'Street_name']
        street_type = test_auto_data_df.loc[person_i,'Street_type']
        suburb = test_auto_data_df.loc[person_i,'Suburb']
        bracket_words = ""
        if "(" in street_name:
            bracket_words = re.findall(r'\((.*?)\)' ,street_name)[0] # extacts all words from the name which are within brackets (e.g. King street (West) will have West extracted)
            street_name = street_name.split("(")[0].strip()
        if "MT " in suburb:
            suburb = suburb.replace("MT", "MOUNT")

        # getting the persons birthdate out as a date object (allows us to get the correct format more easily)
        birthdate = test_auto_data_df.loc[person_i,'DOB']

        # formatting the number of additional drivers to drive the car
        additional_drivers = test_auto_data_df.loc[person_i, "Additional Drivers"]
        if additional_drivers == "No":
            additional_drivers = 0
        else:
            additional_drivers = 1

        # formatting the excess (rounded to the nearest option provided by AA)
        excess = float(test_auto_data_df.loc[person_i,'Excess']) # convert into a floating point value (if it is not already one)
        excess_options = [400, 500, 750, 1000, 1500, 2500] # defines a list of the acceptable 
        # choose the largest excess option for which the customers desired excess is still larger
        excess_index = 0
        while excess >= excess_options[excess_index] and excess_index < 5: # 5 is the index of the largest option, so should not iterate up further if the index has value 5
            excess_index += 1

        # formatting whether or not the car is an automatic
        automatic = test_auto_data_df.loc[person_i,'Gearbox']

        # formatting gearbox info (Number of speeds)
        if " Sp " in  automatic: # if Gearbox starts with 'num Sp ' e.g. (4 Sp ...)
            num_speeds = automatic[0]
        else:
            num_speeds = "" # if the number of speeds is not stated, set it to an empty string

        if "Manual" in automatic: 
            automatic = "Manual" # the way manual is displayed on the AA website
            transmission_type_short = "MAN"
            transmission_type_full = "Manual"
        elif "Automatic" in automatic: # all the different types of automatic transmissions (from the test excel spreadsheeet) and how they are labeled

            automatic = "Auto" # the way they display automatic gearbox on the AA webpage

            if "Sports" in automatic: # if it is a sports automatic
                transmission_type_short = "SPTS AUTO"
                transmission_type_full = "Sports Automatic"
            else: # if its just a regular auotmatic
                transmission_type_short = "AUTO"
                transmission_type_full = "Automatic"
        elif "CVT" in automatic or "Constantly Variable Transmission" in automatic: #CVT/Constantly Variable Transmission (a 'single speed' transmission type (allows smoother transmissions) + is fundementally an automatic))
            automatic = "Other" 
            transmission_type_short = "CVT"
            transmission_type_full = "Constantly Variable Transmission"
        elif "Reduction Gear" in automatic:
            automatic = "Other" 
            transmission_type_short = "Reduction Gear"
            transmission_type_full = "Reduction Gear"
        else:
        # for all other gearboxes (e.g. reduction gear (for electric cars), 
        # or DSG
            automatic = "Other" 
            transmission_type_short = ""
            transmission_type_full = ""

        
        # formatting the engine size
        engine_size = "{}".format(round(float(test_auto_data_df.loc[person_i,'CC']) / 1000, 1))

        # define a dict to store information for a given person and car
        aa_data  = {"Cover_type":test_auto_data_df.loc[person_i,'CoverType'],
                    "AA_member":test_auto_data_df.loc[person_i,'AAMember'],
                    "Registration_number":test_auto_data_df.loc[person_i,'Registration'],
                    "Vehicle_year":test_auto_data_df.loc[person_i,'Vehicle_year'],
                    "Manufacturer":test_auto_data_df.loc[person_i,'Manufacturer'],
                    "Model":str(test_auto_data_df.loc[person_i,'Model']),
                    "Automatic":automatic,
                    "Body_type":test_auto_data_df.loc[person_i,'Body'],
                    "Model_type":model_type,
                    "Model_series":model_series,
                    "Engine_size":engine_size,
                    "Num_speeds":num_speeds.upper(),
                    "Transmission_type_short":transmission_type_short,
                    "Transmission_type_full":transmission_type_full,
                    "Modifications":test_auto_data_df.loc[person_i,'Modifications'],
                    "Finance_purchase":test_auto_data_df.loc[person_i,'FinancePurchase'],
                    "Business_use":test_auto_data_df.loc[person_i,'BusinessUser'],
                    "Street_address": f"{house_number} {street_name} {street_type} {bracket_words}",
                    "Street":f"{street_name} {bracket_words}",
                    "Suburb":suburb.strip(),
                    "Postcode":test_auto_data_df.loc[person_i,'Postcode'],
                    "Birthdate_day":int(birthdate.strftime("%d")),
                    "Birthdate_month":birthdate.strftime("%B"),
                    "Birthdate_year":int(birthdate.strftime("%Y")),
                    "Sex":test_auto_data_df.loc[person_i,'Gender'],
                    "Other_policies":test_auto_data_df.loc[person_i,'OtherPolicies'],
                    "Incidents_3_year":int(test_auto_data_df.loc[person_i,'Incidents_last3years_AA']), # number of at fault crashes in the last 3 years
                    "Current_insurer":test_auto_data_df.loc[person_i, "CurrentInsurer"].upper(),
                    "Additional_drivers":additional_drivers,
                    "Agreed_value":str(round(float(test_auto_data_df.loc[person_i,'AgreedValue']))), # rounds the value to nearest whole number, converts to an integer then into a sting with no dp
                    "Excess_index":excess_index
                    }
        
        
        # adding info on the date and type of incident to the aa_data dictionary ONLY if the person has had an incident within the last 5 years
        if aa_data["Incidents_3_year"] > 0:
            # saving all the incident dates
            for i in range(1, aa_data["Incidents_3_year"] + 1):
                incident_date = test_auto_data_df.loc[person_i,f'Date_of_incident{i}']
                aa_data[f"Incident{i}_date_month"] = incident_date.strftime("%B")
                aa_data[f"Incident{i}_date_year"] = int(incident_date.strftime("%Y"))

            # saving the incident type
            incident_type = test_auto_data_df.loc[person_i,'Type_incident'].lower()
            aa_data["Incident_type"] = "" # initialise "Incident type variable"

            if "not at fault" in incident_type and "no other vehicle involved" in incident_type: # Not at fault - no other vehicle involved
                aa_data["Incident_type"] = "Any claims where no excess was payable" 
            elif "not at fault" in incident_type: # Not at fault - other vehicle involved
                aa_data["Incident_type"] = "Any claims where no excess was payable"
            elif "theft" in incident_type: # At fault - Fire damage or theft
                aa_data["Incident_type"] = "Other Event (eg: Theft, Damaged while parked, Malicious damage)"
            else: # At fault - other vehicle involved
                aa_data["Incident_type"] = "At Fault Event"

        return aa_data


    # scrapes the insurance premium for a single vehicle/person at aa
    def aa_auto_scrape_premium(data):
        def car_info_input_testings(list_element_id, car_spec_value, page_load_wait_time):
            try:
                dropdown = Select(driver.find_element(By.ID, list_element_id)) # find given element as a 'Select' object (easier way to interact with the dropdown list)
                
                # if no option already selected (if dropdown option still on default)
                if dropdown.first_selected_option.text == "Select":
                    dropdown.select_by_value(car_spec_value)
                    time.sleep(page_load_wait_time) # wait for page to load after inputing information

            except:
                print("For", list_element_id, car_spec_value, "wasn't found", end=" - ")
                raise Exception("Exit_to_outer")

        def select_model_variant():
            # scraping these details from the webpage
            car_variant_options = tuple(driver.find_elements(By.XPATH, "//*[@id='vehicleList']/tr"))
            
            # define the specifications list, in the order that we want to use them to filter out incorrect car variant options
            specifications_list = ["Model_type", "Model_series", "Engine_size", "Num_speeds", "Transmission"]

            # iterate through all of the potential car specs that we can use to select the correct drop down option
            for specification in specifications_list:
                # initialise an empty list to store the selected car variants
                selected_car_variants = [] 

                # save the actual value of the specification as a variable (to allow formatting manipulation)
                if specification == "Num_speeds":
                    outer_specification_value = f"{data[specification]} SPEED"
                    inner_specification_value = f"{data[specification]}SP"
                elif specification == "Transmission":
                    outer_specification_value = data["Transmission_type_full"].upper()
                    inner_specification_value = data["Transmission_type_short"].upper()       
                else:
                    outer_specification_value = data[specification].upper()
                    inner_specification_value = data[specification].upper()

                # check all car variants for this specification
                for car_variant in car_variant_options:
                    
                    # Check that all of the known car details are correct (either starts with, ends with, or contains the details as a word in the middle of the text)
                    if car_variant.text.upper().startswith(f"{inner_specification_value} ") or f" {inner_specification_value} " in car_variant.text.upper() or car_variant.text.upper().endswith(f" {inner_specification_value}") \
                        or car_variant.text.upper().startswith(f"{outer_specification_value} ") or f" {outer_specification_value} " in car_variant.text.upper() or car_variant.text.upper().endswith(f" {outer_specification_value}"):

                        # if this car has correct details add it to the select list
                        selected_car_variants.append(car_variant) 


                # checking if we have managed to isolate one option
                if len(selected_car_variants) == 1:
                    # saving the select model variant to the output df
                    partial_car_variant = driver.find_element(By.XPATH, '//*[@id="vehicleDetailSummaryBoxAlt"]/div/div[2]').text
                    aa_output_df.loc[person_i, "AA_selected_car_variant"] = f"{partial_car_variant} {selected_car_variants[0].text}".replace("Found: ","").replace("\n","")

                    # returning the selected car variant
                    return selected_car_variants[0]
                elif len(selected_car_variants) > 1:
                    car_variant_options = tuple(selected_car_variants)
            
            ## choosing the remaining option with the least number of characters
            final_car_variant = car_variant_options[0] # initialising the final variant option to the 1st remaining

            aa_output_df.loc[person_i, "AA_Error_code"] = "Several Car Variant Options Warning"

            # iterating through all other options to find one with least number of characters
            for car_variant in car_variant_options[1:]:

                if len(car_variant.text) < len(final_car_variant.text):
                    final_car_variant = car_variant

            # saving the select model variant to the output df
            partial_car_variant = driver.find_element(By.XPATH, '//*[@id="vehicleDetailSummaryBoxAlt"]/div/div[2]').text

            aa_output_df.loc[person_i, "AA_selected_car_variant"] = f"{partial_car_variant} {final_car_variant.text}".replace("Found: ","").replace("\n"," ")

            return final_car_variant

        def input_agreed_value():
            agreed_value_input = driver.find_element(By.ID, "amountCoveredInput") # find the element to input agreed value into
            agreed_value_input.send_keys(str(10*9)) # input 1 billion (a number that is way too large so we can always scrape the limits)

            # scrapes the aa accepted limits
            Wait.until(EC.visibility_of_element_located( (By.XPATH, "//*[@id='*.errors']/div[2]/ul/li") )) # checks if the error message, that says agreed value is invalid, is present
            agreed_value_limits_list = driver.find_elements(By.XPATH, "//*[@id='*.errors']/div[2]/ul/li/strong[@class='numeric']") # scrapes the min and max values for the agreed value

            limits = [0, 0] # initialise limits list, to store values for the agreed value min and max for a given car

            # pulls out the agreed value limits as integers and saves to limits list
            for i in range(2):
                limits[i] = int(agreed_value_limits_list[i].text.replace(",", "")) 

            aa_output_df.loc[person_i, "AA_agreed_value_minimum"] = limits[0] # save the minimum allowed agreed value
            aa_output_df.loc[person_i, "AA_agreed_value_maximum"] = limits[1] # save the maximum allowed agreed value

            # rounds up/down if the agreed value isnt within the limits
            if int(data["Agreed_value"]) > limits[1]: # if the entered agreed value is greater than the maximum value aa allows
                data["Agreed_value"] = limits[1] # save the new 'adjusted agreed value'
                print("Attempted to input agreed value larger than the maximum", end=" - ")

            elif int(data["Agreed_value"]) < limits[0]: # if the entered agreed value is smaller than the minimum value aa allows
                data["Agreed_value"] = limits[0] # save the new 'adjusted agreed value'
                print("Attempted to input agreed value smaller than the minimum", end=" - ")

            # output the corrected agreed value
            aa_output_df.loc[person_i, "AA_agreed_value"] = data["Agreed_value"]

            # input the amount covered (Agreed Value)
            agreed_value_input = driver.find_element(By.ID, "amountCoveredInput") # find the element to input agreed value into
            agreed_value_input.clear() # clear current values

            time.sleep(1) # wait for page to load

            agreed_value_input.send_keys(data["Agreed_value"]) # input the desired value
        
            time.sleep(2) # wait for page to load

        def enter_registration_number():
            if data["Registration_number"] == "": # if the vehicle registration number is NA then raise an exception (go to except block, skip rest of try)
                raise Exception("Registration_NA")
            else:
                driver.find_element(By.ID, "vehicleRegistrationNumberNz").send_keys(data["Registration_number"]) # input registration number
                driver.find_element(By.ID, "vehicleRegistrationSearchButtonNz").click() # click check button

                time.sleep(1.5) # wait for page to load

                # attempt to find the for car summary pop down (if present then we can continue)
                Wait.until(EC.visibility_of_element_located( (By.ID,  "vehicleDetailSummaryBoxAlt")))
        
        def enter_car_details_manually():

            # if the vehicle registration number is NA then we need to click this button (else if the registration number is just invalid we dont)
            if data["Registration_number"] == "" or not use_registration_number: 
                Wait10.until(EC.element_to_be_clickable((By.ID, "modelSelector-button")) ).click() # click Model Selector button

            try:
                # find year of manufacture list and select the correct year
                dropdown = Select(driver.find_element(By.ID, "vehicleYearOfManufactureList"))
                dropdown.select_by_value(str(data["Vehicle_year"]))

                time.sleep(1) # wait for page to load

                # find car make (manufacturer) list and select the correct manufacturer
                dropdown = Select(driver.find_element(By.ID, "vehicleMakeList"))
                dropdown.select_by_value(data["Manufacturer"].upper())

                time.sleep(2) # wait for page to load

                # checking if the car model is already input, if not already input selects correct model
                car_info_input_testings("vehicleModelList", data["Model"].upper(), 2) 

                # checking if the car transmission is already input, if not already input selects correct transmission type (auto, manual, or other)
                car_info_input_testings("vehicleTransmissionList", data["Automatic"][0], 1)  # only the 1st letter needed

                # checking if the car body type is already input, if not already input selects correct body type
                car_info_input_testings("vehicleBodyTypeList", data["Body_type"], 0) 
            except:
                print(f"Couldn't find {data["Vehicle_year"]} {data["Manufacturer"]} {data["Model"]} with body type {data["Body_type"]} and {data["Automatic"]} transmission", end=" -- ")
                return None

            driver.find_element(By.ID, "findcar").click()  # click find your car button



        # Open the webpage
        driver.get("https://online.aainsurance.co.nz/motor/pub/aainzquote?insuranceType=car")

        # choose cover type
        cover_type_lower = data["Cover_type"].lower()
        if "comprehensive" in cover_type_lower: # if the cover type is comprehensive
            Wait10.until(EC.element_to_be_clickable( (By.XPATH, "/html/body/div[4]/main/div/div[2]/form/fieldset[1]/div/div/label[1]/span") )).click() # click comprehensive cover type button
        elif ("third party" in cover_type_lower) and ("fire" in cover_type_lower) and ("theft" in cover_type_lower): # if cover type is 'third party fire and theft'
            Wait10.until(EC.element_to_be_clickable( (By.XPATH, "//*[@id='quote']/fieldset[1]/div/div/label[3]/span") )).click() # click third party fire and theft cover type button
        else: # otherwise assume the insurance cover type is third party
            Wait10.until(EC.element_to_be_clickable( (By.XPATH, "//*[@id='quote']/fieldset[1]/div/div/label[2]/span") )).click() # click third party cover type button

        # select whether the individual is an AA member
        if data["AA_member"] == "Yes":
            Wait.until(EC.element_to_be_clickable( (By.XPATH, "//*[@id='aaMembershipDetails.aaMemberButtons']/label[1]/span") ) ).click()
        else:
            Wait.until(EC.element_to_be_clickable( (By.XPATH, "//*[@id='aaMembershipDetails.aaMemberButtons']/label[2]/span") ) ).click()

        
        # attempt to input the car registration number (if it both provided and valid)
        try:
            if use_registration_number:
                enter_registration_number()
            else:
                raise Exception("Not Using Registration Number!")
        except: # if the registration is invalid or not provided, then need to enter car details manually
            enter_car_details_manually()
        
        # check if we need to select a model variant
        try:

            time.sleep(1) # wait for page to load

            Wait.until(EC.visibility_of_element_located( (By.ID, "vehicleList-wrapper") )) # checks/ waits until if the pop down to select the variant is visable

            time.sleep(2) # wait for page to load

            # select the correct model variant
            selected_model_variant_element = select_model_variant()

            # click the selected model variant
            selected_model_variant_element.click()


        except exceptions.TimeoutException: # if we dont need to select a model variant then continue on
            selected_car_variant = driver.find_element(By.XPATH, '//*[@id="vehicleDetailSummaryBoxAlt"]/div/div[2]').text

            # saving the select model variant to the output df
            aa_output_df.loc[person_i, "AA_selected_car_variant"] = selected_car_variant.replace("Found: ","").replace("\n"," ")
            pass
            
        # click button to move to car features page
        time.sleep(1) 
        Wait.until(EC.element_to_be_clickable( (By.ID, "_eventId_submit") )).click()


        # click button to state whether or not the car has any modifications
        try:
            if data["Modifications"] == "No":
                Wait10.until(EC.element_to_be_clickable( (By.XPATH, "//*[@id='accessoriesAndModificationsButtons']/label[2]/span") )).click() # click "No" modifications button
            else:
                Wait10.until(EC.element_to_be_clickable( (By.XPATH, "//*[@id='accessoriesAndModificationsButtons']/label[1]/span") )).click() # click "Yes" modifications button
                raise Exception("Not insurable: Modifications")
        except exceptions.TimeoutException:
            try:
                Wait.until(EC.presence_of_element_located((By.ID, "jeopardy")))
                print("We need a bit more information to progress with your quote (requires calling in)", end= " -- ")
                return "Tech-error page"
            except exceptions.TimeoutException:
                print("Unknown issue (from modifactions code section))", end=" -- ")
                return None
        

        # click button to move to car details page
        driver.find_element(By.ID, "_eventId_submit").click()

        # select whether the car was purchased on finance
        if data["Finance_purchase"].lower() == "no":
            Wait10.until(EC.element_to_be_clickable( (By.XPATH, "//*[@id='vehicleFinance.financedButtons']/label[2]/span") )).click() # click "No" Finance" button
        else:
            Wait10.until(EC.element_to_be_clickable( (By.XPATH, "//*[@id='vehicleFinance.financedButtons']/label[1]/span") )).click() # click "Yes" Finance" button

        # select whether the car is for business or private use
        if data["Business_use"].lower() == "yes":
            driver.find_element(By.XPATH, "//*[@id='vehicleUse.vehiclePrimaryUseButtons']/label[2]/span").click() # click 'business' button
        else:
            driver.find_element(By.XPATH, "//*[@id='vehicleUse.vehiclePrimaryUseButtons']/label[1]/span").click() # click 'private' button

        ## input the address the car is usually kept overnight at
        # inputing postcode + suburb
        try:
            time.sleep(1) # wait a bit to let the page load
            driver.find_element(By.ID, "address.suburbPostcodeRegionCity").send_keys(data["Postcode"]) # input the postcode
            time.sleep(3) # wait a bit for the page to load
            Wait.until(EC.element_to_be_clickable( (By.XPATH, f"//*[@id='quote']/fieldset[3]/div[1]/div[1]/ul/li[contains(text(), '{data["Suburb"]}')]") )).click() # find the pop down option with the correct suburb

        except exceptions.TimeoutException:
            Wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='quote']/fieldset[3]/div[1]/div[1]/ul/li[1]"))).click() # if the above doesn't work, just select the first pop down option
        finally: # input street address
            Wait.until(EC.element_to_be_clickable((By.ID, "address.streetAddress")) ).send_keys(data["Street_address"]) # inputting street address

        aa_output_df.loc[person_i, "AA_selected_address"] = f"{data["Street_address"]}, {data["Suburb"]}, {data["Postcode"]}"

        # click button to move to driver details page
        driver.find_element(By.ID, "_eventId_submit").click()

        # if the pop up to further clarify address appears
        try:
            time.sleep(1) # wait for page to load
            Wait.until_not(EC.visibility_of_element_located((By.XPATH, "//*[@id='suggestedAddresses-container']/div[2]/div"))) # checking if pop up, indicating that more info on address is needed, does not appear

        except exceptions.TimeoutException: # go here only if the pop up does appear
            try:
                # finding the options that has the correct street name and postcode
                address_element = driver.find_element(By.XPATH, f"//*[@id='suggestedAddresses-container']/div[2]/div/label/span[contains(text(), '{data["Street"].lower().title()}') and contains(text(), '{data["Postcode"]}')]")

                # write the selected address to the output dataframe
                aa_output_df.loc[person_i, "AA_selected_address"] = address_element.text

                address_element.click() # clicking the address that has the correct street name and postcode

            except exceptions.NoSuchElementException:
                print(f"Couldn't find address, {data["Street_address"]}, {data["Suburb"]} with postcode {data["Postcode"]}", end=" - ")
                return None
            
            driver.find_element(By.ID, "_eventId_submit").click() # click button to move to driver details page
        
        time.sleep(1) # wait for page to load


        # inputing main drivers date of birth
        Wait10.until(EC.element_to_be_clickable( (By.ID, "mainDriver.dateOfBirth-day") )).click() # open DOB day drop down
        driver.find_element(By.XPATH, "//*[@id='mainDriver.dateOfBirth-day']//*[text()='{}']".format(data["Birthdate_day"])).click() # select main driver DOB day
        Wait.until(EC.element_to_be_clickable( (By.ID, "mainDriver.dateOfBirth-month") )).click() # open DOB month drop down
        driver.find_element(By.XPATH, "//*[@id='mainDriver.dateOfBirth-month']//*[text()='{}']".format(data["Birthdate_month"])).click() # select main driver DOB month
        Wait.until(EC.element_to_be_clickable( (By.ID, "mainDriver.dateOfBirth-year") )).click() # open DOB year drop down
        driver.find_element(By.XPATH, "//*[@id='mainDriver.dateOfBirth-year']//*[text()='{}']".format(data["Birthdate_year"])).click() # select main driver DOB year

        # select gender of main driver
        if data["Sex"].lower() == "male":
            driver.find_element(By.XPATH, "//*[@id='mainDriver.driverGenderButtons']/label[1]/span").click() # clicking the 'Male' button
        else: #is Female
            driver.find_element(By.XPATH, "//*[@id='mainDriver.driverGenderButtons']/label[2]/span").click() # clicking the 'Female' button

        # select whether or not the person has any current policies with AA (If yes gives a multi policy discount)
        if data["Other_policies"].lower() == "no": # they don't have other policies with AA
            driver.find_element(By.XPATH, "//*[@id='existingPoliciesButtons']/label[2]/span").click() # click the button saying that we have no current policies with AA
        else: # they have other policies with AA
            driver.find_element(By.XPATH, "//*[@id='mainDriver.driverGenderButtons']/label[1]/span").click() # click the button saying that we have a current policy with AA


        # select the individuals current insurer
        driver.find_element(By.ID, "previousInsurerList").click() # open the drop down for the previous insurers
        Wait.until(EC.element_to_be_clickable( (By.XPATH, f"//*[@id='allPreviousInsurerOptionGroup']/option[contains(text(),'{data["Current_insurer"]}')]") )).click() # click the correct 'previous insurer'
        
        # click the button saying how many accidents you have been in in last 3 years
        driver.find_element(By.XPATH, f"//*[@id='mainDriverNumberOfAccidentsOccurrencesButtons']/label[{data['Incidents_3_year']+1}]/span").click() # click button to say 1 "incident" in last 3 years

        # input the details for all of the recent incidents (if it is 0, then it goes right past)
        for i in range(0, data['Incidents_3_year']):
            driver.find_element(By.ID, f"mainDriver.accidentTheftClaimOccurrenceList[{i}].occurrenceType.accidentTheftClaimOccurrenceType").click() # click button to open type of occurrence pop down
            driver.find_element(By.XPATH, f"//*[@id='mainDriver.accidentTheftClaimOccurrenceList[{i}].occurrenceType.accidentTheftClaimOccurrenceType']/option[text() ='{data["Incident_type"]}']").click() # select the type of occurrence
            
            # input the approximate date it occured
            driver.find_element(By.ID, f"mainDriver.accidentTheftClaimOccurrenceList[{i}].monthOfOccurrence.month").click() # open the month dropdown
            driver.find_element(By.XPATH, f"//*[@id='mainDriver.accidentTheftClaimOccurrenceList[{i}].monthOfOccurrence.month']/option[text() ='{data[f"Incident{i+1}_date_month"]}']").click() # month selection
            driver.find_element(By.ID, f"mainDriver.accidentTheftClaimOccurrenceList[{i}].yearOfOccurrence.year").click() # open the year dropdown
            driver.find_element(By.XPATH, f"//*[@id='mainDriver.accidentTheftClaimOccurrenceList[{i}].yearOfOccurrence.year']/option[text() ='{data[f"Incident{i+1}_date_year"]}']").click() # year selection
        
        # click button to specify how many additional drivers there are
        if data["Additional_drivers"] == 0:
            driver.find_element(By.XPATH, "//*[@id='numberOfAdditionalDriversButtons']/label[1]/span").click()
        else:
            driver.find_element(By.XPATH, "//*[@id='numberOfAdditionalDriversButtons']/label[2]/span").click()
        
        # click the "Get my quote" button
        driver.find_element(By.ID, "_eventId_submit").click()

        # check to see if we are on the page that says 'Unfortunately we can't offer you insurance at this time.' 
        try:
            Wait.until(EC.presence_of_element_located((By.ID, "jeopardy")))
            print("Unfortunately we can't offer you insurance at this time.", end= " -- ")
            return "Tech-error page"
        except exceptions.TimeoutException:
            pass
        

        try:
            Wait.until(EC.presence_of_element_located( (By.ID, "techError") )) # if we went to the 'tech error' page (Says  "Sorry our online service is temporarily unavailable" everytime for this car/person)
            print("We need a bit more information to progress with your quote (requires calling in)", end=" - ")
            return "Tech-error page" # we exit this option if the error page comes up (we cannot scrape for this person/car)
        except exceptions.TimeoutException:
            pass

        # calls a function to input the agreed value
        input_agreed_value()

        # clicks the persons desired excess level
        Wait.until(EC.element_to_be_clickable( (By.XPATH, "//*[@id='excessContainer']/label[{}]".format(data["Excess_index"])) )).click()
        
        # scrape the monthly premium
        driver.find_element(By.ID, "payMonthlySelected").click() # click on montly 'tab'
        time.sleep(1) # wait for page to load
        monthly_premium = driver.find_element(By.XPATH, "//*[@id='monthlyPremiumId']/span").text # get text for the monthly premium amount

        # scrape the yearly premium
        driver.find_element(By.ID, "payAnnuallySelected").click() # click on yearly 'tab'
        time.sleep(1) # wait for page to load
        yearly_premium = driver.find_element(By.XPATH, "//*[@id='yearlyPremiumId']/span").text # get text for the yearly premium amount
        
        # reformatting the montly and yearly premiums into integers
        monthly_premium, yearly_premium = convert_money_str_to_int(monthly_premium, cents=True), convert_money_str_to_int(yearly_premium, cents=True)

        # returning the monthly/yearly premium and the adjusted agreed value
        return monthly_premium, yearly_premium



    # get time of start of each iteration
    start_time = time.time()

    try:
        # scrapes the insurance premium for a single vehicle and person at aa
        aa_auto_premium = aa_auto_scrape_premium(aa_auto_data_format(person_i)) 
        if aa_auto_premium != None and aa_auto_premium != "Tech-error page": # if an actual result is returned

            # print the scraping results
            print( aa_auto_premium[0],  aa_auto_premium[1], end =" -- ")

            # save the scraped premiums to the output dataset
            aa_output_df.loc[person_i, "AA_monthly_premium"] = aa_auto_premium[0] # monthly
            aa_output_df.loc[person_i, "AA_yearly_premium"] = aa_auto_premium[1] # yearly
        elif aa_auto_premium == "Tech-error page":
            aa_output_df.loc[person_i, "AA_Error_code"] = "Webiste Does Not Quote For This Car Variant/ Person"
        else:
            aa_output_df.loc[person_i, "AA_Error_code"] = "Unknown Error"

    except:
        try: # checks if the reason our code failed is because the 'we need more information' pop up appeareds
            Wait.until(EC.visibility_of_element_located( (By.XPATH, "//*[@id='ui-id-3' and text() = 'We need more information']") ) )
            print("Need more information", end= " -- ")
            aa_output_df.loc[person_i, "AA_Error_code"] = "Webiste Does Not Quote For This Car Variant/ Person"
        except exceptions.TimeoutException:
            print("Unknown Error!!", end= " -- ")
            aa_output_df.loc[person_i, "AA_Error_code"] = "Unknown Error"

    end_time = time.time() # get time of end of each iteration
    print("Elapsed time:", round(end_time - start_time,2)) # print out the length of time taken


def auto_scape_all():
    # performing all data reading in and preprocessing
    dataset_preprocess()

    ## reading in variables from the standard input that the 'parent' process passes in
    input_indexes = input()
    input_indexes = input_indexes.replace("[", "").replace("]", "").split(",")

    # saving whether or not to use car registration number while scraping?
    global use_registration_number
    if input_indexes[0] == "Y":
        use_registration_number = True
    else:
        use_registration_number = False

    # save the start index and the number of cars in the dataset as a variable
    input_indexes = list(map(int, input_indexes[1:]))



    # loop through all cars in test spreadsheet
    for person_i in input_indexes: 

        print(f"{person_i}: AA: ", end = "") # print out the iteration number

        # set for this person, the PolicyStartDate to todays date
        test_auto_data_df.loc[person_i, "PolicyStartDate"] = datetime.strftime(date.today(), "%d/%m/%Y")

        # run on the ith car/person
        aa_auto_scrape(person_i)

        # delete all cookies to reset the page
        try:
            driver.delete_all_cookies()
        except exceptions.TimeoutException: # if we timeout while trying to reset the cookies

                print("\n\nNew Webdriver window\n")
                driver.quit() # quit this current driver
                load_webdriver() # open a new webdriver session
    

    export_auto_dataset(input_indexes)




def main():
    # setup the chromedriver options
    options = webdriver.ChromeOptions()
    options.add_experimental_option('excludeSwitches', ['enable-logging'])

    # loads chromedriver
    load_webdriver()

    # scrape all of the insurance premiums for the given cars from aa
    auto_scape_all()

    # Close the browser window
    driver.quit()

main()
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
def export_auto_dataset(input_indexes):
    try:
        # set the column "Sample Number" to be the index column
        global tower_output_df # making it so that this 'reset' of the index is done to the globally stored data
        tower_output_df.set_index("Sample Number", drop=True, inplace=True)

        # read in the output dataset
        insurance_premium_web_scraping_Tower_df = pd.read_csv(f"{parent_dir}\\Individual-company_data-files\\tower_scraped_auto_premiums.csv")
    except FileNotFoundError:
        # if there exists no df already, then we just export what we have scraped
        auto_dataset_for_export = tower_output_df.iloc[input_indexes] # get the given number of lines from the start
    else:
        # set the column "Sample Number" to be the index column so the dataframe combine on "Sample Number"
        insurance_premium_web_scraping_Tower_df.set_index("Sample Number", drop=True, inplace=True)

        # combine with the newly scraped data (anywhere there is an overlap, the newer (just scraped) data overwrites the older data)
        auto_dataset_for_export = tower_output_df.iloc[input_indexes].combine_first(insurance_premium_web_scraping_Tower_df)
    finally:
        # export the dataset
        auto_dataset_for_export.to_csv(f"{parent_dir}\\Individual-company_data-files\\tower_scraped_auto_premiums.csv")

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

# performing the data reading in and preprocessing
def dataset_preprocess():
    # read in the data
    global test_auto_data_df
    test_auto_data_df = pd.read_csv(test_auto_data_csv, dtype={"Postcode":"int"})

    # sets all values of the policy start date to be today's date
    for key in test_auto_data_df:
        test_auto_data_df['PolicyStartDate'] = datetime.strftime(date.today(), "%d/%m/%Y")
    
    # converts the two values which should be dates into dates
    test_auto_data_df['DOB'] = pd.to_datetime(test_auto_data_df['DOB'], format='%Y-%m-%d')
    test_auto_data_df['Date_of_incident'] = pd.to_datetime(test_auto_data_df['Date_of_incident'], format='%Y-%m-%d')

    # pads out the front of postcodes with zeroes (as excel removes leading zeros)
    test_auto_data_df['Postcode'] = test_auto_data_df['Postcode'].apply(postcode_reformat)
    
    # creates a new dataframe to save the scraped info
    global tower_output_df
    tower_output_df = test_auto_data_df.loc[:, ["Sample Number", "PolicyStartDate"]]
    tower_output_df["Tower_agreed_value"] = test_auto_data_df["AgreedValue"].to_string(index=False).strip().split()
    tower_output_df["Tower_monthly_premium"] = ["-1"] * len(test_auto_data_df)
    tower_output_df["Tower_yearly_premium"] = ["-1"] * len(test_auto_data_df)
    tower_output_df["Tower_agreed_value_minimum"] = [-1] * len(test_auto_data_df)
    tower_output_df["Tower_agreed_value_maximum"] = [-1] * len(test_auto_data_df)
    tower_output_df["Tower_Error_code"] = ["No Error"] * len(test_auto_data_df)

def load_webdriver():
    # loads chromedriver
    global driver # defines driver as a global variable
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))

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


# defining a function that will scrape all of the tower cars
def tower_auto_scrape(person_i):
# defining a function which take the information from the spreadsheet and formats it so it can be used to scrape premium from tower website
    def tower_auto_data_format(person_i):
        # saving manufacturer as a variable
        manufacturer = str(test_auto_data_df.loc[person_i,'Manufacturer']).title()

        # formatting model type
        model_type = test_auto_data_df.loc[person_i,'Type']
        if pd.isna(model_type):
            model_type = ""
        elif "XL Dual Cab" in model_type:
            model_type = model_type.replace("Dual", "Double")

        # formatting model series
        model_series = test_auto_data_df.loc[person_i,'Series']

        # setting NA values to an empty string
        if pd.isna(model_series):
            model_series = ""
        else:
            model_series = str(model_series)

        # getting the street address
        street_name = test_auto_data_df.loc[person_i,'Street_name']
        street_type = test_auto_data_df.loc[person_i,'Street_type']
        suburb = test_auto_data_df.loc[person_i,'Suburb']
        if "(" in street_name:
            street_name = street_name.split("(")[0].strip()
        if "MT " in suburb:
            suburb = suburb.replace("MT", "MOUNT")

        # formatting unit number
        unit_number = test_auto_data_df.loc[person_i,'Unit_number']
        if pd.isna(unit_number): # if there is no unit number (is na), set the unit number varibale to an empty string
            unit_number = ""

        # formatting engine size
        engine_size = f"{round(test_auto_data_df.loc[person_i,'CC']/1000, 1)}"
        if engine_size == "0.0": # if there is no cubic centimetres engine measurement for car (is electric)
            engine_size = ""

        # formatting car model type (for when the label is just C)
        model = test_auto_data_df.loc[person_i,'Model']
        if model == "C":
            model += str(engine_size) # add on the number of '10 times litres' in the engine
        
        ## formatting gearbox info
        automatic = test_auto_data_df.loc[person_i,'Gearbox']

        # formatting gearbox info (Number of speeds)
        if " Sp " in  automatic: # if Gearbox starts with 'num Sp ' e.g. (4 Sp ...)
            num_speeds = automatic[:4].replace(" ", "").lower()
        else:
            num_speeds = "" # if the number of speeds is not stated, set it to an empty string

        # formatting gearbox info (what transmission type)
        if "Constantly Variable Transmission" in automatic or "CVT" in automatic:
            automatic = "CVT"  
        elif "Manual" in automatic: 
            automatic = "Man" 
        elif "Automatic" in automatic or "DSG": # all the different types of automatic transmissions (from the test excel spreadsheeet) and how they are labeled
            automatic = "Auto" 
        else:
            automatic = "Other" # for all other gearboxes (e.g. reduction gear in electric)


        # getting the persons birthdate out as a date object (allows us to get the correct format more easily)
        birthdate = test_auto_data_df.loc[person_i,'DOB']

        # formatting the excess (rounded to the nearest option provided)
        excess = float(test_auto_data_df.loc[person_i,'Excess']) # convert into a floating point value (if it is not already one)
        excess_options = [400, 500, 750, 1000] # defines a list of the acceptable 
        # choose the largest excess option for which the customers desired excess is still larger
        excess_index = 0
        while excess > excess_options[excess_index] and excess_index < 3: # 3 is the index of the largest option, so should not iterate up further if the index has value 3
            excess_index += 1

        # define a dict to store information for a given person and car
        tower_data = {"Registration_number":test_auto_data_df.loc[person_i,'Registration'],
                    "Manufacturer":manufacturer,
                    "Model":str(model).title(),
                    "Model_type":model_type,
                    "Model_series":model_series,
                    "Vehicle_year":test_auto_data_df.loc[person_i,'Vehicle_year'],
                    "Body_type":str(test_auto_data_df.loc[person_i,'Body']).title(),
                    "Engine_size":engine_size,
                    "Num_speeds":num_speeds,
                    "Automatic":automatic,
                    "Business_use":test_auto_data_df.loc[person_i,'BusinessUser'],
                    "Unit":unit_number,
                    "Street_number":test_auto_data_df.loc[person_i,'Street_number'],
                    "Street_name":f"{street_name} {street_type}".title(),
                    "Suburb":test_auto_data_df.loc[person_i,'Suburb'].strip().title(),
                    "Postcode":test_auto_data_df.loc[person_i,'Postcode'],
                    "Birthdate_day":int(birthdate.strftime("%d")),
                    "Birthdate_month":birthdate.strftime("%m"),
                    "Birthdate_year":int(birthdate.strftime("%Y")),
                    "Sex":test_auto_data_df.loc[person_i,'Gender'],
                    "Incidents_3_year":test_auto_data_df.loc[person_i,'Incidents_last3years_AA'],
                    "Exclude_under_25":test_auto_data_df.loc[person_i,'ExcludeUnder25'],
                    "Cover_type":test_auto_data_df.loc[person_i,'CoverType'],
                    "Agreed_value":int(round(test_auto_data_df.loc[person_i,'AgreedValue'])), # rounds the value to nearest whole number then converts to an integer
                    "Excess_index":excess_index,
                    "Modifications":test_auto_data_df.loc[person_i,'Modifications'],
                    "Immobiliser":test_auto_data_df.loc[person_i,'Immobiliser_alarm'], 
                    "Finance_purchase":test_auto_data_df.loc[person_i,'FinancePurchase'],
                    "Policy_start_date":str(test_auto_data_df.loc[person_i, "PolicyStartDate"])
                    }
        
        # adding info on the date and type of incident to the tower_data dictionary ONLY if the person has had an incident within the last 5 years
        incident_date = test_auto_data_df.loc[person_i,'Date_of_incident']
        if tower_data["Incidents_3_year"] == "Yes":
            tower_data["Incident_year"] = int(incident_date.strftime("%Y"))

            incident_type = test_auto_data_df.loc[person_i,'Type_incident'].lower() # initialising the variable incident type

            # choosing what Incident_type_index is
            if "theft" in incident_type: # “At fault – Fire damage or Theft”
                tower_data["Incident_type_index"] = 6 # "Theft"
                tower_data["Incident_excess_paid"] = "Yes"
            elif "at fault" in incident_type: #at fault - other vehicle involved"
                tower_data["Incident_type_index"] = 2 # "Collision"
                tower_data["Incident_excess_paid"] = "Yes"
            elif "no other vehicle involved" in incident_type: # “Not at fault – no other vehicle involved”
                tower_data["Incident_type_index"] = 1 # "Broken windscreen"
                tower_data["Incident_excess_paid"] = "No"
            else: # “Not at fault – other vehicle involved”
                tower_data["Incident_type_index"] = 2 # "Collision"
                tower_data["Incident_excess_paid"] = "No"

        # returns the dict object containing all the formatted data
        return tower_data

    def tower_auto_scrape_premium(data):
        def handle_business_use_button():
            # Open the webpage
            driver.get("https://my.tower.co.nz/quote/car/page1")

            # wait for page to load
            time.sleep(3)

            # selects whether or not the car is used for business
            if data["Business_use"] == "Yes":
                Wait10.until(EC.element_to_be_clickable( (By.ID, "btnvehicleUsedForBusiness-0") )).click() # clicks "Yes" business use button
            elif data["Business_use"] == "No":
                Wait10.until(EC.element_to_be_clickable( (By.ID, "btnvehicleUsedForBusiness-1") )).click() # clicks "No" business use button
            else:
                Wait10.until(EC.element_to_be_clickable( (By.ID, "btnvehicleUsedForBusiness-2") )).click() # clicks "Sometimes" business use button

        def select_model_variant():
            # scraping these details from the webpage
            car_variant_options = tuple(driver.find_elements(By.XPATH, '//*[@id="carVehicleTypes-menu-list"]/li'))

            # define the specifications list, in the order that we want to use them to filter out incorrect car variant options
            specifications_list = ["Model_series", "Engine_size", "Num_speeds", "Automatic"]

            # iterate through all of the potential car specs that we can use to select the correct drop down option
            for specification in specifications_list:
                # initialise an empty list to store the selected car variants
                selected_car_variants = [] 

                # save the actual value of the specification as a variable (to allow formatting manipulation depending on what the spec is)
                if specification == "Engine_size":
                    specification_value = f"{data[specification]}i"
                else:
                    specification_value = data[specification]

                # check all car variants for this specification
                for car_variant in car_variant_options:
                    # Check that all of the known car details are correct (either starts with, ends with, or contains the details as a word in the middle of the text)
                    if car_variant.text.startswith(f"{specification_value} ") or f" {specification_value} " in car_variant.text or car_variant.text.endswith(f" {specification_value}"):

                        # if this car has correct details add it to the select list
                        selected_car_variants.append(car_variant) 

                # checking if we have managed to isolate one option
                if len(selected_car_variants) == 1:
                    return selected_car_variants[0]
                elif len(selected_car_variants) > 1:
                    car_variant_options = tuple(selected_car_variants)
            
            ## choosing the remaining option with the least number of characters
            final_car_variant = car_variant_options[0] # initialising the final variant option to the 1st remaining
            print("Unable to fully narrow down", end=" - ")
            tower_output_df.loc[person_i, "Tower_Error_code"] = "Several Car Variant Options Warning"



            # iterating through all other options to find one with least number of characters
            for car_variant in car_variant_options[1:]:
                
                if len(car_variant.text) < len(final_car_variant.text):
                    final_car_variant = car_variant
            
            return final_car_variant


        # attempt to click business use button
        try: # if error go to except
            handle_business_use_button()
        except:
            # if we cannot click the business use button, close and reopen the page
            driver.quit() # quit this current driver
            load_webdriver() # open a new webdriver session

            # Open the webpage
            driver.get("https://my.tower.co.nz/quote/car/page1") 

            try: # if error go to except
                handle_business_use_button()
            except:
                print("Cannot click business use button", end=" -- ")
                return None

                
        
        # clicks the button to reset the input data (need to do this everytime except first time, as the page 'remebers' the previous iteration)
        try:
            Wait.until(EC.presence_of_element_located( (By.ID, "vehicleUsedForBusiness-error-link"))).click()
        except exceptions.TimeoutException:
            pass


        # attempt to input the car registration number (if it both provided and valid)
        try: 
            if pd.isna(data["Registration_number"]): # if the vehicle registration number is NA then raise an exception (go to except block, skip rest of try)
                raise ValueError("Registration_NA")
            else:
                # attempt to input the license plate number. If it doesn't work then raise value error to go enter the car details manually
                driver.find_element(By.ID, "txtLicencePlate").send_keys(data["Registration_number"]) # input registration number
                driver.find_element(By.ID, "btnSubmitLicencePlate").click() # click seach button
                

                time.sleep(1) # wait for page to load
                

                try:
                    Wait.until_not(EC.presence_of_element_located( (By.XPATH, "//*[@id='carLookupError']/div/p[text() = 'No record found']") )) # checks that the carLookupError pop up does NOT appear
                except exceptions.TimeoutException: # if the carLookupError pop up appears
                    raise ValueError("Registration_Invalid")

                try:

                    # find the elements which lists the cars details, after having input the registration number
                    registration_found_car = Wait10.until(EC.presence_of_element_located((By.XPATH, "//*[@id='questionCarLookup']/div[4]/div[@class='car-results']")) )
                    registration_found_car_text = registration_found_car.text

                    # if the car manufacturer, model, body type and year of production are not correct, then enter details manually
                    if not ((str(data["Manufacturer"]) in registration_found_car_text) and (str(data["Model"]) in registration_found_car_text) and (str(data["Body_type"]) in registration_found_car_text) and (str(data["Vehicle_year"]) in registration_found_car_text)):
                        raise ValueError("Car found by registration number is incorrect!")
                    
                    # if the car model type, transmission type, number of speeds and engine size (num of litres) are not correct, then enter details manually
                    extra_registration_found_car_text = registration_found_car.find_element(By.XPATH, ".//p").text
                    if not ( (str(data["Model_type"]) in extra_registration_found_car_text) and (str(data["Automatic"]) in extra_registration_found_car_text) and (str(data["Num_speeds"]) in extra_registration_found_car_text) and (str(data["Engine_size"]) in extra_registration_found_car_text) ):
                        raise Exception("Car found by registration number has several options")

                except: # if the car information automatically found (by entering license plate) is false, then go check if options box is present

                    # try to select correct option 
                    try:
                        # clicking the button which has the correct model type information, as well as the correct Gearbox (Automatic and Num_speeds) as well as Engine_Size
                        Wait.until(EC.element_to_be_clickable( (By.XPATH, f"//*[@id='questionCarLookup']/div[4]/div[2]/fieldset/div/label/span[contains(text(), '{data["Model_type"]}') and " + 
                                                            f"contains(text(), '{data["Automatic"]}') and " + 
                                                            f"contains(text(), '{data["Num_speeds"]}') and " +
                                                            f"contains(text(), '{data["Engine_size"]}')]") )).click()
                    except exceptions.TimeoutException: # if the options box not present, then enter the details manually
                        Wait.until(EC.element_to_be_clickable( (By.ID, "lnkEnterMakeModel") )).click() # Find the button "Enter your car's details" and click it
                        raise ValueError("Car information incorrect")


        except ValueError: # if the registration is invalid or not provided, then need to enter car details manually
            
            time.sleep(1) # wait for page to load

            # Find the button "Enter your car's details" and click it
            if pd.isna(data["Registration_number"]): # only if registration NA do we need to click the button (if the plate number is just invalid, then the drop down automatically opens, so dont need to click this button)
                Wait.until(EC.element_to_be_clickable( (By.ID, "lnkEnterMakeModel") )).click()

            # inputting the car manufacturer
            try:
                # find car manufacturer input box and input the company that manufactures the car
                manufacturer_text_input = driver.find_element(By.ID, "carMakes")
                if manufacturer_text_input.get_attribute("value") == "": # checks the input fields value is currently empty
                    manufacturer_text_input.send_keys(data["Manufacturer"])

                    Wait10.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="carMakes-wrapper"]/div/i'))) # wait for page to load
                    
                    # click the button to select the car manufacturer in the dropdown (i just click the 1st drop down option because I assume this must be the correct one)
                    Wait.until(EC.element_to_be_clickable( (By.XPATH, "//*[@id='carMakes-menu-list']/li/div/div[2]/div") )).click() 
            except exceptions.TimeoutException:
                print(f"CANNOT FIND {data["Manufacturer"]}", end=" -- ")
                return None # return None if can't scrape

            # inputting car model
            try:
                # wait until car model input box is clickable, then input the car model
                car_model_text_input = Wait.until(EC.presence_of_element_located((By.ID, "carModels")))
                if car_model_text_input.get_attribute("value") == "": # checks the input fields value is currently empty
                    car_model_text_input.send_keys(data["Model"]) 

                    Wait10.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="carModels-wrapper"]/div/i'))) # wait for page to load

                    # wait until button which has the correct car model information is clickable, then click (i just click the 1st drop down option because I assume this must be the correct)
                    Wait.until(EC.element_to_be_clickable( (By.XPATH, "//*[@id='carModels-menu-list']/li[1]/div/div[2]/div") )).click() 
            except exceptions.TimeoutException:
                print(f"CANNOT FIND {data["Manufacturer"]} MODEL {data["Model"]}", end=" -- ")
                return None # return None if can't scrape

            # inputting car year
            try:
                car_model_text_input = Wait.until(EC.presence_of_element_located( (By.ID, "carYears") )) # find car year input box
                if car_model_text_input.get_attribute("value") == "": # checks the input fields value is currently empty
                    car_model_text_input.send_keys(str(data["Vehicle_year"])) # inputs the year 

                    Wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="carYears-wrapper"]/div/i'))) # wait for page to load

                    Wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='carYears-menu-list']/li[1]"))).click() # clicking the button which has the correct car year information
            except exceptions.TimeoutException:
                print(f"CANNOT FIND {data["Manufacturer"]} {data["Model"]} FROM YEAR {data["Vehicle_year"]}", end=" -- ")
                return None # return None if can't scrape


            # inputting car body style
            try:
                body_type_text_input = Wait.until(EC.presence_of_element_located((By.ID, "carBodyStyles"))) # find the car body type input box and
                if body_type_text_input.get_attribute("value") == "": # checks the input fields value is currently empty

                    body_type_text_input.send_keys(data["Body_type"]) # inputs the body type

                    Wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="carBodyStyles-wrapper"]/div/i'))) # wait for page to load

                    Wait.until(EC.element_to_be_clickable( (By.XPATH, "//*[@id='carBodyStyles-menu-list']/li[1]/div/div[2]/div") )).click() # clicking the button which has the correct car body style information
            except exceptions.TimeoutException: # if code timeout while waiting for element
                print(f"CANNOT FIND {data["Vehicle_year"]} {data["Manufacturer"]} {data["Model"]} WITH BODY TYPE {data["Body_type"]}", end=" -- ")
                return None # return None if can't scrape

            # inputting car vehicle type
            car_model_type_text_input = Wait.until(EC.presence_of_element_located((By.ID, "carVehicleTypes"))) # find the model type input box and then input the model type
            if car_model_type_text_input.get_attribute("value") == "": # checks the input fields value is currently empty
                car_model_type_text_input.send_keys(data["Model_type"]) # inputs the body type

                time.sleep(10000)

                # select the correct model variant
                selected_model_variant_element = select_model_variant()

                # click the selected model variant
                selected_model_variant_element.click()

        # enters whether the car has an immobiliser, if needed
        try:
            if data["Immobiliser"] == "Yes":
                Wait.until(EC.element_to_be_clickable( (By.ID, "btnvehicleAlarm-0") )).click()
            else: # if No Immobiliser alarm
                Wait.until(EC.element_to_be_clickable( (By.ID, "btnvehicleAlarm-1") )).click()
        except exceptions.TimeoutException:
            pass

        time.sleep(1) # wait for page to process information

        ## input the address the car is usually kept overnight at
        # inputing postcode + suburb
        Wait.until(EC.element_to_be_clickable( (By.ID, "lnkManualAddress") )).click() # click this button to enter the address manually
        try:
            Wait.until(EC.presence_of_element_located((By.ID, "txtAddress-street-number")) ).send_keys(data["Street_number"]) # input the street number
            Wait.until(EC.presence_of_element_located((By.ID, "txtAddress-flat-number")) ).send_keys(data["Unit"]) # input the flat number (Unit number)
            Wait.until(EC.presence_of_element_located((By.ID, "txtAddress-suburb-city-postcode")) ).send_keys(data["Street_name"] + ", " + data["Suburb"]) # input the street name and suburb

            # find the pop down option with the correct suburb and postcode then select it 
            Wait10.until(EC.element_to_be_clickable( (By.XPATH, f"//ul[@id='txtAddress-suburb-city-postcode-menu-list']/*/div/div[2]/div[contains(text(), '{data["Suburb"]}') and contains(text(), '{data["Postcode"]}')]"))).click()
        except exceptions.TimeoutException:
            try:
                driver.find_element(By.XPATH, f"//ul[@id='txtAddress-suburb-city-postcode-menu-list']/*/div/div[2]/div[contains(text(), '{data["Postcode"]}')]").click() # if the above doesn't work, just select an option with the correct postcode
            except exceptions.NoSuchElementException:
                raise Exception("Address Invalid")
            

        # enter the persons 'name' entering placeholder pseudonym either Jane or John Doe
        if data["Sex"] == "MALE":
            first_name = "John"
        else:
            first_name = "Jane"
        
        Wait.until(EC.presence_of_element_located( (By.ID, "txtDriver-0-firstName") )).send_keys(first_name)
        Wait.until(EC.presence_of_element_located( (By.ID, "txtDriver-0-lastName") )).send_keys("Doe")


        # enter the main drivers date of birth
        Wait.until(EC.presence_of_element_located( (By.ID, "driverDob-0-day") )).send_keys(data["Birthdate_day"]) # input day
        Wait.until(EC.presence_of_element_located( (By.ID, "driverDob-0-month") )).send_keys(data["Birthdate_month"]) # input month
        Wait.until(EC.presence_of_element_located( (By.ID, "driverDob-0-year") )).send_keys(data["Birthdate_year"]) # input year

        try:
            time.sleep(0.5) # wait for page to load
            Wait.until_not(EC.presence_of_element_located( (By.ID, "driver-dob-error") )) # checks if the person is old enough for tower to accept insuring them for the given car
        except: # if the warning is present, then this car/ person combo cannot be insured (the person is 'too yonug' for the given car)
            print("We will only cover your vehicle when it is being driven by people aged 25 or over", end=" - ")
            return "Doesn't Cover"

        # select gender of main driver
        if data["Sex"] == "Male":
            driver.find_element(By.ID, "btndriverGender-0-0").click() # clicking the 'Male' button
        else: #is Female
            driver.find_element(By.ID, "btndriverGender-0-1").click() # clicking the 'Female' button


        # input if there have been any indicents
        if data["Incidents_3_year"] == "Yes":
            driver.find_element(By.ID, "btndriverVehicleLoss-0-0").click() # clicks button saying that you have had an incident

            driver.find_element(By.ID, "driverVehicleLossReason-0-0-toggle").click() # open incident type dropdown

            # Incident_type_index == ...; 1: "Broken windscreen", 2: "Collision", 6: "Theft"
            Wait.until(EC.presence_of_element_located( (By.XPATH, f"//ul[@id='driverVehicleLossReason-0-0-menu-options']/li[{data["Incident_type_index"]}]") )).click() # find the correct incident then select it

            time.sleep(1) # wait for page to load


            # selects the correct year the indicent occured
            year_id_index = datetime.now().year - data["Incident_year"] # gets the index for the id of the correct year element
            Wait.until(EC.presence_of_element_located( (By.ID, f"btndriverVehicleLossReasonWhen-0-0-{year_id_index}") )).click()


            # if the incident type is a collision, damage while parked, or other causes
            if data["Incident_type_index"] == 2:

                # if they had to pay an excess (if they made a claim)
                if data["Incident_excess_paid"] == "Yes":
                    driver.find_element(By.ID, "btndriverVehicleLossExcess-0-0-0").click() # click yes button
                else:
                    driver.find_element(By.ID, "btndriverVehicleLossExcess-0-0-1").click() # click no button
        else:
            driver.find_element(By.ID, "btndriverVehicleLoss-0-1").click() # clicks button saying that you have had no incidents

        # choose whether to exclude under all under 25 year old drivers from driving the car
        try:
            if data["Exclude_under_25"] == "Yes":\
                driver.find_element(By.ID, "btnexcludeUnder25-0").click() # clicks yes button for exlcude under 25
            else: # select 'No'
                driver.find_element(By.ID, "btnexcludeUnder25-1").click() # clicks no button for exlcude under 25
        except:
            pass
        
        # press button to aknowledge the extra $1000 excess
        try:
            Wait.until(EC.element_to_be_clickable((By.ID, "btnClose"))).click()
            print("Extra $1000 Excess", end=" - ")
            time.sleep(10000)
        except exceptions.TimeoutException:
            pass

        try:
            # try to click button 'Next: Customise' to move onto next page
            Wait.until(EC.element_to_be_clickable( (By.ID, "btnSubmitPage") )).click()
        except exceptions.TimeoutException:
            # accept the privacy policy
            Wait.until(EC.element_to_be_clickable( (By.ID, "privacyPolicy-label") )).click()

            # THEN click button 'Next: Customise' to move onto next page
            Wait.until(EC.element_to_be_clickable( (By.ID, "btnSubmitPage") )).click()

        time.sleep(2.5) # wait for page to load


        # choose cover type
        cover_type_title = data["Cover_type"].lower()
        if "comprehensive" in cover_type_title: # if the cover type is comprehensive
            time.sleep(2) # comprehensive cover type is by default selected (so we just wait a bit for page to load)
        elif ("third party" in cover_type_title) and ("fire" in cover_type_title) and ("theft" in cover_type_title): # if cover type is 'third party fire and theft'
            Wait10.until(EC.element_to_be_clickable( (By.ID, "buttonCoverType-1") )).click() # click third party fire and theft cover type button
        else: # otherwise assume the insurance cover type is third party
            Wait10.until(EC.element_to_be_clickable( (By.ID, "buttonCoverType-2") )).click() # click third party cover type button


        time.sleep(4) # wait for the page to load


        ## input the amount covered (Agreed Value)
        # scrapes the max and min values
        min_value = convert_money_str_to_int(Wait10.until(EC.presence_of_element_located( (By.XPATH, "//*[@id='agreedValueNew']/div[1]/div[1]/div[2]") )).text) # get the min agreed value
        max_value = convert_money_str_to_int(Wait10.until(EC.presence_of_element_located( (By.XPATH, "//*[@id='agreedValueNew']/div[1]/div[3]/div[2]") )).text) # get the max agreed value
        
        tower_output_df.loc[person_i, "Tower_agreed_value_minimum"] = min_value # save the minimum allowed agreed value
        tower_output_df.loc[person_i, "Tower_agreed_value_maximum"] = max_value # save the maximum allowed agreed value

        # check if our attempted agreed value is valid. if not, round up/down to the min/max value
        if data["Agreed_value"] > max_value:
            data["Agreed_value"] = max_value
            print("Attempted to input agreed value larger than the maximum", end=" - ")
        elif data["Agreed_value"] < min_value:
            data["Agreed_value"] = min_value
            print("Attempted to input agreed value smaller than the minimum", end=" - ")

        # output the corrected agreed value
        tower_output_df.loc[person_i, "Tower_agreed_value"] = data["Agreed_value"]

        # inputs the agreed value input the input field (after making sure its valid)
        agreed_value_input = driver.find_element(By.ID, "agreedValueNewSliderField") # find the input field for the agreed value
        agreed_value_input.send_keys(Keys.CONTROL, "a") # select all current value
        agreed_value_input.send_keys(data["Agreed_value"]) # input the desired value, writing over the (selected) current value
        driver.find_element(By.ID, "agreedValueNewSliderBtn").click() # click the 'Update agreed value' button


        time.sleep(1) # wait for page to load


        # input the persons desired level of excess
        Wait10.until(EC.element_to_be_clickable( (By.ID, f"btnexcess-{data["Excess_index"]}"))).click()

        # click button to state whether or not the car has any 'major' modifications
        if data["Modifications"] == "No":
            no_button = Wait10.until(EC.element_to_be_clickable( (By.XPATH, "//*[@id='btnaccessoriesOrModificationsDeclined-1']/div[2]/div") )) # find "No" modifications button
            driver.execute_script("arguments[0].scrollIntoView();", no_button) # scroll down until "No" modifications button is on screen (is needed to prevent the click from being intercepted)
            no_button.click() # click "No" modifications button
        else:
            yes_button = Wait10.until(EC.element_to_be_clickable( (By.XPATH, "//*[@id='btnaccessoriesOrModificationsDeclined-0']/div[2]/div") ))# find "Yes" modifications button
            driver.execute_script("arguments[0].scrollIntoView();", yes_button) # scroll down until "Yes" modifications button is on screen (is needed to prevent the click from being intercepted)
            yes_button.click() # click "Yes" modifications button
            raise Exception("Not insurable: Modifications")

        # click button to state whether or not the car has any 'minor' modifications
        if data["Modifications"] == "No":
            Wait10.until(EC.element_to_be_clickable( (By.ID, "btnaccessoriesOrModifications-1") )).click() # click "No" modifications button
        else:
            Wait10.until(EC.element_to_be_clickable( (By.ID, "btnaccessoriesOrModifications-0") )).click() # click "Yes" modifications button
            raise Exception("Need to know the combined value of the minor modifications")
        


        # click 'Next: Summary' button to move onto the summary page
        next_button = Wait10.until(EC.element_to_be_clickable( (By.ID, "btnSubmitPage")))
        driver.execute_script("arguments[0].scrollIntoView();", next_button) # scroll down until "Next: Summary" button is on screen (is needed to prevent the click from being intercepted)
        Wait10.until(EC.element_to_be_clickable((By.ID, "btnSubmitPage"))).click()

        time.sleep(5)

        # move onto the next page "People"
        Wait10.until(EC.element_to_be_clickable((By.ID, "btnSubmitPage"))).click()

        time.sleep(3) # wait for page to load    
    
        # if a pop telling us 'an additional excess of ... will apply to claims of theft for this vehicle' appears, close it
        try:
            Wait.until(EC.element_to_be_clickable((By.ID, "btnClose"))).click()
        except exceptions.TimeoutException: # if the popup not present then continue
            pass

        # click button to say that the driver has NOT "had their licence suspended or cancelled or had a special condition imposed"
        Wait10.until(EC.element_to_be_clickable((By.ID, "btndriver-0-licence-cancelled-1"))).click()

        # click button to say that the policy will NOT be held by a business or trust
        Wait.until(EC.element_to_be_clickable((By.ID, "btnownedByBusinessOrTrust-1"))).click()

        # enter email address
        email_address = f"{first_name}.Doe@email.com"
        Wait.until(EC.presence_of_element_located((By.ID, "txtDriver-0-email"))).send_keys(email_address)

        # enter phone number
        Wait.until(EC.presence_of_element_located((By.ID, "txtDriver-0-phoneNumbers-0"))).send_keys("022 123 456")

        # click button to go to next page 'Legal'
        Wait.until(EC.element_to_be_clickable((By.ID, "btnSubmitPage"))).click()

        # click button to say I understand the 1st legal information declaration (the 'Yes' button)
        Wait10.until(EC.element_to_be_clickable((By.ID, "btnlegalDeclaration-0"))).click()

        # click button to say I understand the 2nd legal information (important things to call out) declaration (the 'Yes' button)
        Wait.until(EC.element_to_be_clickable((By.ID, "btnexclusions-0"))).click()

        # click button to say I haven't had insurance refused or cancelled within the last 7 years (the 'No' button)
        Wait.until(EC.element_to_be_clickable((By.ID, "btninsuranceHistory-1"))).click()

        # click button to say I haven't had a claim declined or policy avoided in the last 7 years (the 'No' button)
        Wait.until(EC.element_to_be_clickable((By.ID, "btnclaimsDeclined-1"))).click()

        # click button to say I have not been convicted of Fraud, Arson, Bugulary, Wilfull damage, sexual offences, or drugs conviction within the last 7 years (the 'No' button)
        Wait.until(EC.element_to_be_clickable((By.ID, "btncriminalHistory-1"))).click()

        # select whether the car was purchased on finance
        if data["Finance_purchase"].upper() == "NO":
            Wait10.until(EC.element_to_be_clickable( (By.ID, "btnmoneyOwed-1") )).click() # click "No" Finance" button
        else:
            Wait10.until(EC.element_to_be_clickable( (By.ID, "btnmoneyOwed-0") )).click() # click "Yes" Finance" button
            Wait.until(EC.presence_of_element_located((By.ID, "financialInterestedParty-0-financialInterestedParty-financial-institution-search"))).send_keys("Kiwibank") # enter the finance provider as kiwibank
            Wait.until(EC.element_to_be_clickable((By.XPATH, '//ul[@id="financialInterestedParty-0-financialInterestedParty-financial-institution-search-menu-list"]/li'))).click() # select 'Kiwibank Limited' as finance provider

        # click button to go to the next page 'Finalise'
        Wait.until(EC.element_to_be_clickable((By.ID, "btnSubmitPage"))).click()

        # click button to say the person would not like to link an airpoints account (Click 'No' button)
        Wait.until(EC.element_to_be_clickable((By.ID, "btnairpointsIsMember-1"))).click()

        # input the desired start date
        start_date_input_element = Wait.until(EC.presence_of_element_located((By.ID, "policyStartDatePicker")))
        start_date_input_element.send_keys(Keys.CONTROL + "a")
        start_date_input_element.send_keys(data["Policy_start_date"])

        # click button to move to next page ('Payment')
        Wait10.until(EC.element_to_be_clickable((By.ID, "btnSubmitPage"))).click()

        # scrape the monthly and yearly premiums
        monthly_premium = convert_money_str_to_int(Wait10.until(EC.presence_of_element_located( (By.ID, "btnpaymentFrequency-1"))).text, cents=True) # scrape the monthy premium and convert into an integer
        yearly_premium = convert_money_str_to_int(Wait10.until(EC.presence_of_element_located( (By.ID, "btnpaymentFrequency-2"))).text, cents=True) # scrape the yearly premium and convert into an integer


        # return the scraped premiums
        return round(monthly_premium, 2), round(yearly_premium, 2)


    # get time of start of each iteration
    start_time = time.time() 

    # run on the ith car/person
    try:
        # scrapes the insurance premium for a single vehicle and person
        tower_auto_premium = tower_auto_scrape_premium(tower_auto_data_format(person_i)) 
        if tower_auto_premium != None and tower_auto_premium != "Doesn't Cover": # if an actual result is returned

            # print the scraping results
            print(tower_auto_premium[0], tower_auto_premium[1], end =" -- ")

            # save the scraped premiums to the output dataset
            tower_output_df.loc[person_i, "Tower_monthly_premium"] = tower_auto_premium[0] # monthly
            tower_output_df.loc[person_i, "Tower_yearly_premium"] = tower_auto_premium[1] # yearly
        elif tower_auto_premium == "Doesn't Cover":
            tower_output_df.loc[person_i, "Tower_Error_code"] = "Webiste Does Not Quote For This Car Variant"
        else:
            tower_output_df.loc[person_i, "Tower_Error_code"] = "Unknown Error"

    except:
        try: # checks if the reason our code failed is because the 'we need more information' pop up appeareds
            Wait.until(EC.visibility_of_element_located( (By.XPATH, "//*[@id='ui-id-3' and text() = 'We need more information']") ) )
            print("Need more information", end= " -- ")
            tower_output_df.loc[person_i, "Tower_Error_code"] = "Webiste Does Not Quote For This Car Variant"
        except exceptions.TimeoutException:
            print("Unknown Error!!", end= " -- ")
            tower_output_df.loc[person_i, "Tower_Error_code"] = "Unknown Error"

    end_time = time.time() # get time of end of each iteration
    print("Elapsed time:", round(end_time - start_time, 2)) # print out the length of time taken
    



def auto_scape_all():
    # call the function to read in a preprocess the data
    dataset_preprocess()

    # save the start index and the number of cars in the dataset as a variable (reading it from the standard input that the 'parent' process passes in)
    input_indexes = input()
    input_indexes = input_indexes.replace("[", "").replace("]", "").split(",")
    input_indexes = list(map(int, input_indexes))

    # loop through all cars in test spreadsheet
    for person_i in input_indexes: 

        print(f"{person_i}: Tower: ", end = "") # print out the iteration number

        # run on the ith car/person
        tower_auto_scrape(person_i)

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

    # scrape all of the insurance premiums for the given cars-person combinations
    auto_scape_all()

    # Close the browser window
    driver.quit()

main()
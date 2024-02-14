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
test_auto_data_xlsx = f"{parent_dir}\\test_auto_data1.xlsx" # defines the path to the input dataset



"""
-------------------------
Useful functions
"""
def export_auto_dataset(num_datalines_to_export = len(test_auto_data_xlsx)):
    auto_dataset_for_export = ami_output_df.head(num_datalines_to_export) # get the given number of lines from the start
    auto_dataset_for_export.to_csv(f"{parent_dir}\\Individual-company_data-files\\ami_scraped_auto_premiums.csv", index=False)


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

# defining a function that will scrape all of the ami cars
def ami_auto_scrape(person_i):

    # defining a function which take the information from the spreadsheet and formats it so it can be used to scrape premium from ami website
    def ami_auto_data_format(person_i):
        # formatting street name and type into the correct format
        street_name = test_auto_data.loc[person_i,'Street_name']
        street_type = test_auto_data.loc[person_i,'Street_type']
        suburb = test_auto_data.loc[person_i,'Suburb'].strip()
        if "(" in street_name:
            street_name = street_name.split("(")[0].strip()
        if "MT " in suburb:
            suburb = suburb.replace("MT", "MOUNT")

        # formatting car model type
        model = test_auto_data.loc[person_i,'Model']
        if model == "C":
            model += str(int(math.ceil(test_auto_data.loc[person_i,'CC']/100))) # add on the number of '10 times litres' in the engine

        # getting the persons birthdate out as a date object (allows us to get the correct format more easily)
        birthdate = test_auto_data.loc[person_i,'DOB']

        # formatting drivers licence type for input
        drivers_license_type = ""
        if test_auto_data.loc[person_i,'Licence'] == "NEW ZEALAND FULL LICENCE":
            drivers_license_type = "NZ Full" 
        elif test_auto_data.loc[person_i,'Licence'] == "RESTRICTED LICENCE":
            drivers_license_type = "NZ Restricted" 
        elif test_auto_data.loc[person_i,'Licence'] == "LEARNERS LICENCE":
            drivers_license_type = "NZ Learners" 
        else: # for for generic 'International' (non-NZ) licence
            drivers_license_type = "International / Other overseas licence" 

        # formatting the number of years the person had had their drivers licence
        drivers_license_years = test_auto_data.loc[person_i,'License_years_TOWER']
        if drivers_license_years < 1:
            drivers_license_years = "Less than a year"
        elif drivers_license_years == 1:
            drivers_license_years = "{} year".format(drivers_license_years)
        elif drivers_license_years >= 5:
            drivers_license_years = "5 years or more"
        else: # for for generic 'International' (non-NZ) licence
            drivers_license_years = "{} years".format(drivers_license_years)
        
        # formatting the excess (rounded to the nearest option provided)
        excess = float(test_auto_data.loc[person_i,'Excess']) # convert into a floating point value (if it is not already one)
        excess_options = [100, 400, 500, 1000, 2000] # defines a list of the acceptable 
        # choose the smallest excess option  which is larger than (or equal to) the customers desired excess level
        excess_index = 0
        while excess > excess_options[excess_index] and excess_index < 4 : # 4 is the index of the largest option, so should not iterate up further if the index has value 4
            excess_index += 1
        excess_index += 1 # add on extra to the value of the excess index (as the option buttons for choosing the excess start at 1, not 0)
        # define a dict to store information for a given person and car for ami
        ami_data = {"Registration_number":test_auto_data.loc[person_i,'Registration'],
                    "Manufacturer":test_auto_data.loc[person_i,'Manufacturer'],
                    "Model":model,
                    "Model_type":test_auto_data.loc[person_i,'Type'],
                    "Vehicle_year":test_auto_data.loc[person_i,'Vehicle_year'],
                    "Body_type":test_auto_data.loc[person_i,'Body'].upper(),
                    "Engine_size":"{}cc".format(int(test_auto_data.loc[person_i,'CC'])),
                    "Immobiliser":test_auto_data.loc[person_i,'Immobiliser_alarm'],
                    "Business_use":test_auto_data.loc[person_i,'BusinessUser'],
                    "Unit":test_auto_data.loc[person_i,'Unit_number'],
                    "Street_number":test_auto_data.loc[person_i,'Street_number'],
                    "Street_name":street_name + " " + street_type,
                    "Suburb":suburb,
                    "Postcode":test_auto_data.loc[person_i,'Postcode'],
                    "Birthdate_day":int(birthdate.strftime("%d")),
                    "Birthdate_month":birthdate.strftime("%B"),
                    "Birthdate_year":int(birthdate.strftime("%Y")),
                    "Sex":test_auto_data.loc[person_i,'Gender'],
                    "Drivers_license_type":drivers_license_type,
                    "Drivers_license_years":drivers_license_years, # years since driver got their learners licence
                    "Incidents_5_year":test_auto_data.loc[person_i,'Incidents_last5years_AMISTATE'],
                    "NZ_citizen_or_resident":test_auto_data.loc[person_i,'NZ_citizen_or_resident'],
                    "1_year_Visa":test_auto_data.loc[person_i,'Visa_at_least_1_year'],
                    "Agreed_value":test_auto_data.loc[person_i, "AgreedValue"],
                    "Excess_index":str(excess_index)
                    }
        
        # adding info on the date and type of incident to the ami_data dictionary ONLY if the person has had an incident within the last 5 years
        incident_date = test_auto_data.loc[person_i,'Date_of_incident']
        if ami_data["Incidents_5_year"] == "Yes":
            ami_data["Incident_date_month"] = incident_date.strftime("%B")
            ami_data["Incident_date_year"] = int(incident_date.strftime("%Y"))
            incident_type = test_auto_data.loc[person_i,'Type_incident'].split("-")[0].strip()
            if incident_type == "Not at fault":
                ami_data["Incident_type"] = "Not At Fault Accident"
            else:
                ami_data["Incident_type"] = "At Fault Accident"

        # returns the dict object containing all the formatted data
        return ami_data

    # scrapes the insurance premium for a single vehicle and person at ami
    def ami_auto_scrape_premium(data):

        # defining a function to select the correct model variant
        def select_model_variant():
            # scraping these details from the webpage
            car_variant_options = tuple(driver.find_elements(By.XPATH, '//*[@id="searchByMMYResult"]/div[2]/span'))

            # define the specifications list, in the order that we want to use them to filter out incorrect car variant options
            specifications_list = ["Model_type", "Model_series", "Engine_size"]

            # iterate through all of the potential car specs that we can use to select the correct drop down option
            for specification in specifications_list:
                # initialise an empty list to store the selected car variants
                selected_car_variants = [] 

                # save the actual value of the specification as a variable (to allow formatting manipulation)
                if specification == "Num_speeds":
                    specification_value = f"{data[specification]}SP"
                elif specification == "Transmission":
                    specification_value = data["Transmission_type_short"].upper()       
                else:
                    specification_value = data[specification]

                # check all car variants for this specification
                for car_variant in car_variant_options:
                    
                    # Check that all of the known car details are correct (either starts with, ends with, or contains the details as a word in the middle of the text)
                    if car_variant.text.upper().startswith(f"{specification_value} ") or f" {specification_value} " in car_variant.text.upper() or car_variant.text.upper().endswith(f" {specification_value}"):

                        # if this car has correct details add it to the select list
                        selected_car_variants.append(car_variant) 


                # checking if we have managed to isolate one option
                if len(selected_car_variants) == 1:
                    return selected_car_variants[0]
                elif len(selected_car_variants) > 1:
                    car_variant_options = tuple(selected_car_variants)
            
            ## choosing the remaining option with the least number of characters
            final_car_variant = car_variant_options[0] # initialising the final variant option to the 1st remaining
            print("unable to fully narrow down", end=" - ")

            # iterating through all other options to find one with least number of characters
            for car_variant in car_variant_options[1:]:

                #print(car_variant.text) # print out all remaining options, for checking

                if len(car_variant.text) < len(final_car_variant.text):
                    final_car_variant = car_variant
            
            '''
            print() # print just a newline character

            # printing a message to notify what is happened
            print(f"SELECTED: {final_car_variant.text}. ACTUAL: {data["Vehicle_year"]} {data["Manufacturer"]} {data["Model"]} {data["Model_type"]}" + 
                  f"{data["Model_series"]} with body type {data["Body_type"]} and {data["Num_speeds"]} {data["Automatic"]} and {data["Engine_size"]}L engine", end=" - ")
            '''
            return final_car_variant


        # Open the webpage
        driver.get("https://secure.ami.co.nz/css/car/step1")



        # attempt to input the car registration number (if it both provided and valid)
        registration_na = pd.isna(data["Registration_number"])
        if not registration_na: # if there is a registration number provided
            Wait10.until(EC.presence_of_element_located((By.ID, "vehicle_searchRegNo")) ).send_keys(data["Registration_number"]) # input registration
            driver.find_element(By.ID, "ie_regSubmitButton").click() # click submit button

            time.sleep(2.5)

            # attempt to find the 1st option for car pop down (if present then we can continue)
            try: 

                Wait.until(EC.element_to_be_clickable( (By.ID,  "searchedVehicleSpan_0")))
                registration_invalid = False
            except: # if that element is not findable then the registration must have been invalid
                registration_invalid = True
        # is effectively an "else" statement for the above if
        if registration_na or registration_invalid: # if registration invalid or not provided we need to enter car details
            # Find the button "Make, Model, Year" and click it
            Wait10.until(EC.element_to_be_clickable( (By.ID, "ie_MMYPrepareButton") )).click()


            # inputting the car manufacturer
            car_manfacturer_element = driver.find_element(By.ID, "vehicleManufacturer") # find car manufacturer input box
            car_manfacturer_element.click() # open the input box
            time.sleep(1) # wait for page to process information
            car_manfacturer_element.send_keys(data["Manufacturer"]) # input the company that manufactures the car
            time.sleep(1.5) # wait for page to process information
            try:
                Wait.until(EC.element_to_be_clickable( (By.XPATH, "//a[@class='ui-corner-all' and text()='{}']".format(data["Manufacturer"]) ) )).click() # clicking the button which has the correct manufacturer information
            except exceptions.TimeoutException:
                print("CANNOT FIND {manufacturer}".format(manufacturer = data["Manufacturer"]), end=" -- ")
                return None # return None if can't scrape

            # inputting car model
            try:
                Wait.until(EC.element_to_be_clickable((By.ID, "Model"))).click() # wait until car model input box is clickable, then open it
                Wait.until(EC.element_to_be_clickable((By.XPATH, "//div[text()='{}']".format(data["Model"])))).click() # wait until button which has the correct car model information is clickable, then click
            except exceptions.TimeoutException:
                print("CANNOT FIND {manufacturer} MODEL {model}".format(year = data["Vehicle_year"], manufacturer = data["Manufacturer"], model = test_auto_data.loc[person_i,'Model']), end=" -- ")
                return None # return None if can't scrape

            # inputting car year
            try:
                Wait.until_not(lambda x: x.find_element(By.ID, "searchByMMYLoading").is_displayed()) # wait until the "loading element" is not being displayed
                time.sleep(2) 
                driver.find_element(By.ID, "Year").click()
                Wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[6]/div/div[text()='{}']".format(data["Vehicle_year"])))).click() # clicking the button which has the correct car model information
            except exceptions.TimeoutException:
                print("CANNOT FIND {manufacturer} {model} FROM YEAR {year}".format(year = data["Vehicle_year"], manufacturer = data["Manufacturer"], model = test_auto_data.loc[person_i,'Model']), end=" -- ")
                return None # return None if can't scrape

            # inputting car body type
            try:
                Wait.until_not(lambda x: x.find_element(By.ID, "searchByMMYLoading").is_displayed()) # wait until the "loading element" is not being displayed
                time.sleep(1)
                driver.find_element(By.ID, "BodyType").click() # find car BodyType input box and open it
                Wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[7]/div/div[text()='{}']".format(data["Body_type"])))).click() # clicking the button which has the correct car model information
            except exceptions.TimeoutException: # if code timeout while waiting for element
                print("CANNOT FIND {year} {manufacturer} {model} WITH BODY TYPE {body_type}".format(year = data["Vehicle_year"], manufacturer = data["Manufacturer"], model = data["Model"], body_type = data["Body_type"]), end=" -- ")
                return None # return None if can't scrape



            # inputting car engine size
            try:
                Wait.until_not(lambda x: x.find_element(By.ID, "searchByMMYLoading").is_displayed()) # wait until the "loading element" is not being displayed
                driver.find_element(By.ID, "EngineSize").click() # find car BodyType input box and open it
                Wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[8]/div/div[contains(text(), '{}')]".format(data["Engine_size"])))).click() # clicking the button which has the correct car model information
            except exceptions.TimeoutException:
                print("CANNOT FIND {year} {manufacturer} {model} {body_type}, WITH {engine_size}".format(year = data["Vehicle_year"], manufacturer = data["Manufacturer"], model = data["Model"], body_type = data["Body_type"], engine_size = data["Engine_size"]), end=" -- ")
                return None # return None if can't scrape
            time.sleep(1) # wait for page to process information
            
        # select the final vehicle option
        try:
            if pd.isna(data["Model_type"]):
                raise Exception("NA Model_type") # if the model type is NA we raise an exception, thus going to bottom except block (which it just clicks first option)

            # select the correct model variant
            selected_model_variant_element = select_model_variant()

            # click the selected model variant
            selected_model_variant_element.click()

        except: # Selects the 1st option, if a Model type is not specified
            Wait.until(EC.element_to_be_clickable( (By.ID,  "searchedVehicleSpan_0"))).click() # wait until clickable, then click button to select final vehicle option
                
        # selects whether or not the car has an immobiliser
        try: # we 'try' this because the option to select Immobiliser only comes up on some cars (if there are some models of the car which don't)
            if data["Immobiliser"] == "Yes":
                driver.find_element(By.ID, "bHasImmobilizer_true").click() # clicks True button
            else:   
                driver.find_element(By.ID, "bHasImmobilizer_false").click() # clicks False button
        except:
            pass # if the button isn't present we move on

        # selects whether or not the car is used for business
        try:
            if data["Business_use"] == "No":
                driver.find_element(By.ID, "bIsBusinessUse_false").click() # clicks "False" button
            else:
                driver.find_element(By.ID, "bIsBusinessUse_true").click() # clicks "False" button
        except:
            print("Cannot click business use button", end=" -- ")
            return None

        # inputs the address the car is kept at
        driver.find_element(By.ID, "garagingAddress_autoManualRadio").click() # click button to enter address manually
        if not pd.isna(data["Unit"]): driver.find_element(By.ID, "garagingAddress_manualUnitNumber").send_keys(data["Unit"]) # input Unit/Apt IF is applicable
        driver.find_element(By.ID, "garagingAddress_manualStreetNumber").send_keys(data["Street_number"])
        driver.find_element(By.ID, "garagingAddress_manualStreetName").send_keys(data["Street_name"])
        try: # this try block is all just attempting various ways of selecting the final address, either through selecting a pop down from the street, or a pop down from the suburb
            time.sleep(3) # wait for options list to pop up
            driver.find_element(By.XPATH, "//li[@class='ui-menu-item']//a[contains(text(), '{}') or contains(text(),'{}')]".format(data["Suburb"], data["Postcode"])).click()
        except: # if no pop up after inputting the street address, try inputting the suburb
            suburb_entry_element = driver.find_element(By.ID, "garagingAddress_manualSuburb")
            suburb_entry_element.send_keys(data["Suburb"])
            time.sleep(2) # wait for elements on the page to load
            try:
                Wait.until(EC.element_to_be_clickable((By.XPATH, "//li[@class='ui-menu-item']//a[contains(text(), '{}')]".format(data["Postcode"]) )) ).click() # try to find and click any pop down element that contains the postcode
            except exceptions.TimeoutException:
                try: # try entering just the postcode into the suburb
                    suburb_entry_element.clear() # clears the textbox
                    suburb_entry_element.send_keys(data["Postcode"]) # type into the box just the postcode
                    time.sleep(2)
                    driver.find_element(By.XPATH, "//li[@class='ui-menu-item']//a[contains(text(), '{}')]".format(data["Postcode"])).click() # try to find and click any pop down element that contains the postcode
                except:
                    driver.find_element(By.ID, "garagingAddress_manualUnitNumber").click() # click this button to get out of "Suburb/Town" element
                    if driver.find_element(By.ID, "errorSuburbTownPostcode").is_displayed(): # if an error message appears saying that suburb/town not appearing
                            print("Unable to find address {street_number} {street_name} in the suburb of {suburb} with postcode {postcode}".format(street_number = data["Street_number"], street_name = data["Street_name"], suburb = data["Suburb"], postcode = data["Postcode"]), end=" -- ")
                            return None
                    else:
                        raise Exception("Other problem!!!!")
        
        # enter drivers birthdate
        driver.find_element(By.ID, "driverDay_1").send_keys(data["Birthdate_day"]) # input day
        driver.find_element(By.ID, "driverMonth_1").send_keys(data["Birthdate_month"]) # input month
        driver.find_element(By.ID, "driverYear_1").send_keys(data["Birthdate_year"]) # input year

        # select driver sex
        if data["Sex"] == "MALE":
            driver.find_element(By.ID, "male_1").click() # selects male
        else:
            driver.find_element(By.ID, "female_1").click() # selects female

        # enter drivers licence info
        driver.find_element(By.ID, "DriverLicenceType_1").click() # open the drivers license type options box
        driver.find_element(By.XPATH, "//div[text()='{}']".format(data["Drivers_license_type"])).click() # select the drivers license type
        if "International" in data["Drivers_license_type"]: # if international licence
            if data["NZ_citizen_or_resident"] == "Yes": # is the person NZ citizen/ perm resident
                driver.find_element(By.ID, 'prOrCitizen_1').click()
            else:
                driver.find_element(By.ID, 'notPrOrCitizen_1').click()
                if data["1_year_Visa"] == "Yes": # is the visa of the non-perm resident valid for more than one year
                    driver.find_element(By.ID, 'validVisa_1').click()
                else:
                    driver.find_element(By.ID, 'notValidVisa_1').click()
        driver.find_element(By.ID, "DriverYearsOfDriving_1").click() # open years since got learners box
        driver.find_element(By.XPATH, "//div[text()='{}']".format(data["Drivers_license_years"])).click() # select correct years since got learners

        # input if there have been any indicents
        if data["Incidents_5_year"] == "Yes":
            driver.find_element(By.NAME, "driverLoss").click() # clicks button saying that you have had an incident
            driver.find_element(By.ID, "DriverIncidentType_1").click() # opens incident type option box
            driver.find_element(By.XPATH, "//div[text()='{}']".format(data["Incident_type"])).click() # selects the driver incident type

            driver.find_element(By.ID, "DriverIncidentMonth_1").click() # opens incident month option box
            driver.find_element(By.XPATH, "//html//body//div[15]//div//div[text()='{}']".format(data["Incident_date_month"])).click() # selects the driver incident type
            driver.find_element(By.ID, "DriverIncidentYear_1").click() # opens incident year option box
            driver.find_element(By.XPATH, "//html//body//div[16]//div[text()='{}']".format(data["Incident_date_year"])).click() # selects the driver incident type
        else:
            driver.find_element(By.NAME, "driverNoLoss").click() # clicks button saying that you have had no incidents
        
        time.sleep(1) # wait a bit for the page to load

        # click button to get quote 
        Wait.until(EC.element_to_be_clickable((By.ID, "quoteSaveButton"))).click() # wait until button clickable then click

        # check to see if the "Need more information" popup appears. If it does, then exit
        try:
            Wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="getQuoteError153"]/div[2]/div[1]/button/span[1]')))
            print("Need more information!")
            return None
        except exceptions.TimeoutException:
            pass

        ## input the amount covered (Agreed Value)
        # scrapes the max and min values
        min_value = convert_money_str_to_int(Wait10.until(EC.presence_of_element_located( (By.XPATH, "//*[@id='slider']/span[1]") )).text) # get the min agreed value
        max_value = convert_money_str_to_int(Wait10.until(EC.presence_of_element_located( (By.XPATH, "//*[@id='slider']/span[2]") )).text) # get the max agreed value
        
        # check if our attempted agreed value is valid. if not, round up/down to the min/max value
        if data["Agreed_value"] > max_value:
            data["Agreed_value"] = max_value
            adjusted_agreed_value = max_value
            ami_output_df.loc[person_i, "AMI_agreed_value_was_adjusted"] = 1 # save this value to say that the agreed value was adjusted upwards
            print("Attempted to input agreed value larger than the maximum", end=" - ")
        elif data["Agreed_value"] < min_value:
            data["Agreed_value"] = min_value
            adjusted_agreed_value = min_value
            ami_output_df.loc[person_i, "AMI_agreed_value_was_adjusted"] = -1 # save this value to say that the agreed value was adjusted downwards
            print("Attempted to input agreed value smaller than the minimum", end=" - ")


        # inputs the agreed value input the input field (after making sure its valid)
        agreed_value_input = driver.find_element(By.ID, "agreedValueText") # find the input field for the agreed value
        agreed_value_input.send_keys(Keys.CONTROL, "a") # select all current value
        agreed_value_input.send_keys(str(data["Agreed_value"])) # input the desired value, writing over the (selected) current value

        time.sleep(2) # wait for page to load

        # check that the 'something is wrong' popup is not present, if it is closes it
        try:
            Wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='errorRateQuote']/div[2]/div[1]/button")) ).click()
            print("Clicked x")
        except exceptions.TimeoutException:
            pass
        
        # input the persons desired level of excess
        try:
            # adjusting the excess level (if not present then we can assume that we can't adjust the excess)
            Wait.until(EC.presence_of_element_located((By.XPATH, f"//*[@id='optionExcessSlider']/span[3]")) ).click()

            time.sleep(7) # wait for page to update the final premiums

        except exceptions.TimeoutException:
            print("Unchangable excess", end=" -- ")

            time.sleep(3) # wait for page to update the final premiums

        # scrape the premium
        annual_risk_premium = Wait.until(EC.presence_of_element_located((By.ID, "annualRiskPremium")))

        monthly_premium = float(driver.find_element(By.ID, "dollars").text.replace(",", "") + driver.find_element(By.ID, "cents").text)
        yearly_premium = float(annual_risk_premium.text.replace(",", "")[1:])

        # return the scraped premiums
        try:
            return monthly_premium, yearly_premium, adjusted_agreed_value
        except UnboundLocalError: # if no value saved for adjusted_agreed value, then just return None for it
            return monthly_premium, yearly_premium, None


    # get time of start of each iteration
    start_time = time.time() 

    # run on the ith car/person
    try:
        # scrapes the insurance premium for a single vehicle and person
        ami_auto_premium = ami_auto_scrape_premium(ami_auto_data_format(person_i)) 
        if ami_auto_premium != None: # if an actual result is returned

            # print the scraping results
            print( ami_auto_premium[0],  ami_auto_premium[1], end =" -- ")

            # save the scraped premiums to the output dataset
            ami_output_df.loc[person_i, "AMI_monthly_premium"] = ami_auto_premium[0] # monthly
            ami_output_df.loc[person_i, "AMI_yearly_premium"] = ami_auto_premium[1] # yearly

    except:
        try: # checks if the reason our code failed is because the 'we need more information' pop up appeareds
            Wait.until(EC.visibility_of_element_located( (By.XPATH, "//*[@id='ui-id-3' and text() = 'We need more information']") ) )
            print("Need more information", end= " -- ")
        except exceptions.TimeoutException:
            print("Unknown Error!!", end= " -- ")

    end_time = time.time() # get time of end of each iteration
    print("Elapsed time:", round(end_time - start_time,2)) # print out the length of time taken


def auto_scape_all():
    # read in the data
    global test_auto_data
    test_auto_data = pd.read_excel(test_auto_data_xlsx, dtype={"Postcode":"int"})

    # sets all values of the policy start date to be today's date
    for key in test_auto_data:
        test_auto_data['PolicyStartDate'] = datetime.strftime(date.today(), "%d/%m/%Y")

    # pads out the front of postcodes with zeroes (as excel removes leading zeros)
    test_auto_data['Postcode'] = test_auto_data['Postcode'].apply(postcode_reformat) 
    
    # creates a new dataframe to save the scraped info
    global ami_output_df
    ami_output_df = test_auto_data.loc[:, ["Sample Number", "PolicyStartDate"]]
    ami_output_df["AMI_agreed_value"] = test_auto_data.loc[:, "AgreedValue"]
    ami_output_df["AMI_monthly_premium"] = ["-1"] * len(test_auto_data)
    ami_output_df["AMI_yearly_premium"] = ["-1"] * len(test_auto_data)
    ami_output_df["AMI_agreed_value_was_adjusted"] = [0] * len(test_auto_data)

    # save the number of cars in the dataset as a variable (reading it from the standard input that the 'parent' process passes in)
    num_cars = int(input())

    # loop through all cars in test spreadsheet
    for person_i in range(0, num_cars): 

        print(f"{person_i}: AMI: ", end = "") # print out the iteration number

        # run on the ith car/person
        ami_auto_scrape(person_i)

        # delete all cookies to reset the page
        try:
            driver.delete_all_cookies()
        except exceptions.TimeoutException: # if we timeout while trying to reset the cookies

                print("\n\nNew Webdriver window\n")
                driver.quit() # quit this current driver
                load_webdriver() # open a new webdriver session
    

    export_auto_dataset(num_cars)



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
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
    auto_dataset_for_export = aa_output_df.head(num_datalines_to_export) # get the given number of lines from the start
    auto_dataset_for_export.to_csv(f"{parent_dir}\\Individual-company_data-files\\aa_scraped_auto_premiums.csv", index=False)


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

# a function to open the webdriver (chrome simulation)
def load_webdriver():
    # loads chromedriver
    global driver # defines driver as a global variableaa
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

# defining a function that will scrape the premium for a single car from aa
def aa_auto_scrape(person_i):
    # defining a function which take the information from the spreadsheet and formats it so it can be used to scrape premium from aa website
    def aa_auto_data_format(person_i):
        # formatting model type
        model_type = test_auto_data.loc[person_i,'Type']

        # setting NA values to be an empty string
        if pd.isna(model_type):
            model_type = ""
        elif "XL Dual Cab" in model_type:
            model_type = model_type.replace("Dual", "Double").upper()
        else:
            model_type = str(model_type).upper()

        # formatting model series
        model_series = test_auto_data.loc[person_i,'Series']

        # setting NA values to an empty string
        if pd.isna(model_series):
            model_series = ""
        else:
            model_series = str(model_series)


        # getting the street address
        house_number = remove_non_numeric( str(test_auto_data.loc[person_i,'Street_number']) ) # removes all non-numeric characters from the house number (e.g. removes A from 14A)
        street_name = test_auto_data.loc[person_i,'Street_name']
        street_type = test_auto_data.loc[person_i,'Street_type']
        suburb = test_auto_data.loc[person_i,'Suburb']
        bracket_words = ""
        if "(" in street_name:
            bracket_words = re.findall(r'\((.*?)\)' ,street_name)[0] # extacts all words from the name which are within brackets (e.g. King street (West) will have West extracted)
            street_name = street_name.split("(")[0].strip()
        if "MT " in suburb:
            suburb = suburb.replace("MT", "MOUNT")

        # getting the persons birthdate out as a date object (allows us to get the correct format more easily)
        birthdate = test_auto_data.loc[person_i,'DOB']

        # formatting the indicidents the individual has had in the last 3 years into an integer (either 0 or 1)
        if test_auto_data.loc[person_i,'Incidents_last3years_AA'] == "No":
            Incidents_3_year = 0
        else:
            Incidents_3_year = 1

        # formatting the current insurer information
        current_insurer = test_auto_data.loc[person_i, "CurrentInsurer"]
        if current_insurer == "No current insurer":
            current_insurer = "NONE"
        else:
            current_insurer = current_insurer.upper()
        
        # formatting the number of additional drivers to drive the car
        additional_drivers = test_auto_data.loc[person_i, "Additional Drivers"]
        if additional_drivers == "No":
            additional_drivers = 0
        else:
            additional_drivers = 1

        # formatting the excess (rounded to the nearest option provided by AA)
        excess = float(test_auto_data.loc[person_i,'Excess']) # convert into a floating point value (if it is not already one)
        excess_options = [400, 500, 750, 1000, 1500, 2500] # defines a list of the acceptable 
        # choose the largest excess option for which the customers desired excess is still larger
        excess_index = 0
        while excess >= excess_options[excess_index] and excess_index < 5: # 5 is the index of the largest option, so should not iterate up further if the index has value 5
            excess_index += 1

        # formatting whether or not the car is an automatic
        automatic = test_auto_data.loc[person_i,'Gearbox']

        # formatting gearbox info (Number of speeds)
        if " Sp " in  automatic: # if Gearbox starts with 'num Sp ' e.g. (4 Sp ...)
            num_speeds = str(automatic[0])
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
        engine_size = "{}".format(round(test_auto_data.loc[person_i,'CC'] / 1000, 1))

        # define a dict to store information for a given person and car for ami
        aa_data  = {"Cover_type":test_auto_data.loc[person_i,'CoverType'],
                    "AA_member":test_auto_data.loc[person_i,'AAMember'],
                    "Registration_number":test_auto_data.loc[person_i,'Registration'],
                    "Vehicle_year":test_auto_data.loc[person_i,'Vehicle_year'],
                    "Manufacturer":test_auto_data.loc[person_i,'Manufacturer'],
                    "Model":str(test_auto_data.loc[person_i,'Model']),
                    "Automatic":automatic,
                    "Body_type":test_auto_data.loc[person_i,'Body'],
                    "Model_type":model_type,
                    "Model_series":model_series,
                    "Engine_size":engine_size,
                    "Num_speeds":num_speeds.upper(),
                    "Transmission_type_short":transmission_type_short,
                    "Transmission_type_full":transmission_type_full,
                    "Modifications":test_auto_data.loc[person_i,'Modifications'],
                    "Finance_purchase":test_auto_data.loc[person_i,'FinancePurchase'],
                    "Business_use":test_auto_data.loc[person_i,'BusinessUser'],
                    "Street_address": f"{house_number} {street_name} {street_type} {bracket_words}",
                    "Street":f"{street_name} {bracket_words}",
                    "Suburb":suburb.strip(),
                    "Postcode":test_auto_data.loc[person_i,'Postcode'],
                    "Birthdate_day":int(birthdate.strftime("%d")),
                    "Birthdate_month":birthdate.strftime("%B"),
                    "Birthdate_year":int(birthdate.strftime("%Y")),
                    "Sex":test_auto_data.loc[person_i,'Gender'],
                    "Incidents_3_year":Incidents_3_year,
                    "Current_insurer":current_insurer,
                    "Additional_drivers":additional_drivers,
                    "Agreed_value":str(int(round(test_auto_data.loc[person_i,'AgreedValue']))), # rounds the value to nearest whole number, converts to an integer then into a sting with no dp
                    "Excess_index":excess_index
                    }
        
        
        # adding info on the date and type of incident to the ami_data dictionary ONLY if the person has had an incident within the last 5 years
        incident_date = test_auto_data.loc[person_i,'Date_of_incident']
        if aa_data["Incidents_3_year"] == 1:
            aa_data["Incident_date_month"] = incident_date.strftime("%B")
            aa_data["Incident_date_year"] = int(incident_date.strftime("%Y"))
            incident_type = test_auto_data.loc[person_i,'Type_incident'].lower()
            aa_data["Incident_type"] = "" # initialise "Incident type variable"
            if "not at fault" in incident_type or "no other vehicle involved" in incident_type: # if the accident was not at fault and the accident did not involve another vehicle
                aa_data["Incident_type"] = "Any claims where no excess was payable" # mapped 'Not at fault -other vehicle involved' to this
            elif "not at fault" in incident_type: # if the accident was not at fault and the accident involved another vehicle
                aa_data["Incident_type"] = "Any claims where no excess was payable" # mapped 'Not at fault -other vehicle involved' to this
            else: # if the accident was at fault
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
                    return selected_car_variants[0]
                elif len(selected_car_variants) > 1:
                    car_variant_options = tuple(selected_car_variants)
            
            ## choosing the remaining option with the least number of characters
            final_car_variant = car_variant_options[0] # initialising the final variant option to the 1st remaining
            
            #print(f"\n{car_variant_options[0].text}") # print for debugging

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
            if pd.isna(data["Registration_number"]): # if the vehicle registration number is NA then raise an exception (go to except block, skip rest of try)
                raise Exception("Registration_NA")
            else:
                driver.find_element(By.ID, "vehicleRegistrationNumberNz").send_keys(data["Registration_number"]) # input registration number
                driver.find_element(By.ID, "vehicleRegistrationSearchButtonNz").click() # click check button

                time.sleep(1.5) # wait for page to load

                # attempt to find the for car summary pop down (if present then we can continue)
                Wait.until(EC.visibility_of_element_located( (By.ID,  "vehicleDetailSummaryBoxAlt")))
        except: # if the registration is invalid or not provided, then need to enter car details manually
                
                if pd.isna(data["Registration_number"]): # if the vehicle registration number is NA then we need to click this button (else if the registration number is just invalid we dont)
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

        
        # check if we need to select a model variant
        try:

            time.sleep(1) # wait for page to load

            Wait.until(EC.visibility_of_element_located( (By.ID, "vehicleList-wrapper") )) # checks/ waits until if the pop down to select the variant is visable

            time.sleep(2) # wait for page to load

            # select the correct model variant
            selected_model_variant_element = select_model_variant()


            # if we couldnt find a singular correct model variant (is relic from testing)
            if selected_model_variant_element == None:

                # wait for debugging purposes
                input_text = ""
                while input_text != "go":
                    time.sleep(1) # wait for debugging
                    input_text = input("Enter 'go' to continue: ")


                print(end=" -- ") # print for formatting the time elapsed
                return None

            # click the selected model variant
            selected_model_variant_element.click()

        except exceptions.TimeoutException: # if we dont need to select a model variant then continue on
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
                return None
            except exceptions.TimeoutException:
                print("Unknown issue (from modifactions code section))", end=" -- ")
                return None
        

        # click button to move to car details page
        driver.find_element(By.ID, "_eventId_submit").click()

        # select whether the car was purchased on finance
        if data["Finance_purchase"].upper() == "NO":
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

        # click button to move to driver details page
        driver.find_element(By.ID, "_eventId_submit").click()

        # if the pop up to further clarify address appears
        try:
            time.sleep(1) # wait for page to load
            Wait.until_not(EC.visibility_of_element_located((By.XPATH, "//*[@id='suggestedAddresses-container']/div[2]/div"))) # checking if pop up, indicating that more info on address is needed, does not appear

        except exceptions.TimeoutException: # go here only if the pop up does appear
            try:
                driver.find_element(By.XPATH, "//*[@id='suggestedAddresses-container']/div[2]/div/label/span[contains(text(), '{}') and contains(text(), '{}')]".format(data["Street"].lower().title(),
                                                                                                                                                                      data["Postcode"])).click() # clicking the address that has the correct street name and postcode
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
        if data["Sex"] == "Male":
            driver.find_element(By.XPATH, "//*[@id='mainDriver.driverGenderButtons']/label[1]/span").click() # clicking the 'Male' button
        else: #is Female
            driver.find_element(By.XPATH, "//*[@id='mainDriver.driverGenderButtons']/label[2]/span").click() # clicking the 'Female' button

        # click the button saying that we have no current policies with AA (we assume this to get a blank slate)
        driver.find_element(By.XPATH, "//*[@id='existingPoliciesButtons']/label[2]/span").click()

        # select the individuals current insurer
        driver.find_element(By.ID, "previousInsurerList").click() # open the drop down for the previous insurers
        Wait.until(EC.element_to_be_clickable( (By.XPATH, "//*[@id='allPreviousInsurerOptionGroup']/option[contains(text(),'{}')]".format(data["Current_insurer"])) )).click() # click the correct 'previous insurer'

        # click the button saying how many accidents you have been in in last 3 years
        if data["Incidents_3_year"] == 0: # if the person has NOT been in an incident in the last 3 years
            driver.find_element(By.XPATH, "//*[@id='mainDriverNumberOfAccidentsOccurrencesButtons']/label[1]/span").click()
        else: # if the person has been in an incident in the last 3 years
            driver.find_element(By.XPATH, "//*[@id='mainDriverNumberOfAccidentsOccurrencesButtons']/label[2]/span").click() # click button to say 1 "incident" in last 3 years
            driver.find_element(By.ID, "mainDriver.accidentTheftClaimOccurrenceList[0].occurrenceType.accidentTheftClaimOccurrenceType").click() # click button to open type of occurrence pop down
            driver.find_element(By.XPATH, "//*[@id='mainDriver.accidentTheftClaimOccurrenceList[0].occurrenceType.accidentTheftClaimOccurrenceType']/option[text() ='{}']".format(data["Incident_type"])).click() # select the type of occurrence
            
            # input the approximate date it occured
            driver.find_element(By.ID, "mainDriver.accidentTheftClaimOccurrenceList[0].monthOfOccurrence.month").click() # open the month dropdown
            driver.find_element(By.XPATH, "//*[@id='mainDriver.accidentTheftClaimOccurrenceList[0].monthOfOccurrence.month']/option[text() ='{}']".format(data["Incident_date_month"])).click() # month selection
            driver.find_element(By.ID, "mainDriver.accidentTheftClaimOccurrenceList[0].yearOfOccurrence.year").click() # open the year dropdown
            driver.find_element(By.XPATH, "//*[@id='mainDriver.accidentTheftClaimOccurrenceList[0].yearOfOccurrence.year']/option[text() ='{}']".format(data["Incident_date_year"])).click() # year selection
            


        # click button to specify how many additional drivers there are
        if data["Additional_drivers"] == 0:
            driver.find_element(By.XPATH, "//*[@id='numberOfAdditionalDriversButtons']/label[1]/span").click()
        else:
            driver.find_element(By.XPATH, "//*[@id='numberOfAdditionalDriversButtons']/label[2]/span").click()
        
        # click the "Get my quote" button
        driver.find_element(By.ID, "_eventId_submit").click()

        try:
            Wait.until(EC.presence_of_element_located( (By.ID, "techError") )) # if we went to the 'tech error' page (Says  "Sorry our online service is temporarily unavailable" everytime for this car/person)
            print("Tech-error page", end=" - ")
            return None # we exit this option if the error page comes up (we cannot scrape for this person/car)
        except exceptions.TimeoutException:
            pass

        # input the amount covered (Agreed Value)
        agreed_value_input = driver.find_element(By.ID, "amountCoveredInput") # find the element to input agreed value into
        agreed_value_input.clear() # clear current values

        time.sleep(1) # wait for page to load

        agreed_value_input.send_keys(data["Agreed_value"]) # input the desired value

        time.sleep(1) # wait for page to load

        # checks the agreed value is within the aa accepted limits, rounds up/down if it isnt within the limits
        try:
            Wait.until(EC.visibility_of_element_located( (By.XPATH, "//*[@id='*.errors']/div[2]/ul/li") )) # checks if the error message, that says agreed value is invalid, is present
            agreed_value_limits_list = driver.find_elements(By.XPATH, "//*[@id='*.errors']/div[2]/ul/li/strong[@class='numeric']") # scrapes the min and max values for the agreed value

            limits = [0, 0] # initialise limits list, to store values for the agreed value min and max for a given car

            # pulls out the agreed value limits as integers and saves to limits list
            for i in range(2):
                limits[i] = int(agreed_value_limits_list[i].text.replace(",", "")) 
            
            agreed_value_input.clear() # clears the input space to allow us to replace
            
            time.sleep(1) # wait for page to load

            if int(data["Agreed_value"]) > limits[1]: # if the entered agreed value is greater than the maximum value aa allows
                agreed_value_input.send_keys(limits[1]) # input the maximum allowed value
                adjusted_agreed_value = limits[1] # save the new 'adjusted agreed value'
                print("Attempted to input agreed value larger than the maximum", end=" - ")

            elif int(data["Agreed_value"]) < limits[0]: # if the entered agreed value is smaller than the minimum value aa allows
                agreed_value_input.send_keys(limits[0]) # input the minimum allowed value
                adjusted_agreed_value = limits[0] # save the new 'adjusted agreed value'
                print("Attempted to input agreed value smaller than the minimum", end=" - ")

        except exceptions.TimeoutException:
            pass
        
        time.sleep(2) # wait for page to load

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
        try:
            return monthly_premium, yearly_premium, adjusted_agreed_value
        except UnboundLocalError: # if no value saved for adjusted_agreed value, then just return None for it
            return monthly_premium, yearly_premium, None
    




    # get time of start of each iteration
    start_time = time.time()

    try:
        # scrapes the insurance premium for a single vehicle and person at aa
        aa_auto_premium = aa_auto_scrape_premium(aa_auto_data_format(person_i)) 
        if aa_auto_premium != None: # if an actual result is returned

            # print the scraping results
            print( aa_auto_premium[0],  aa_auto_premium[1], end =" -- ")

            # save the scraped premiums to the output dataset
            aa_output_df.loc[person_i, "AA_monthly_premium"] = aa_auto_premium[0] # monthly
            aa_output_df.loc[person_i, "AA_yearly_premium"] = aa_auto_premium[1] # yearly

            # if we adjusted the agreed_value, then save to the output dataset
            if aa_auto_premium[2] != None:
                aa_output_df.loc[person_i, "AgreedValue"] = aa_auto_premium[2] # the adjusted agreed value
                aa_output_df.loc[person_i, "Agreed_value_was_adjusted"] = 1 # save this value to say that the agreed value was adjusted


    except:
        #try: # checks if the reason our code failed is because the 'we need more information' pop up appeareds
        Wait.until(EC.visibility_of_element_located( (By.XPATH, "//*[@id='ui-id-3' and text() = 'We need more information']") ) )
        print("Need more information", end= " -- ")
        #except exceptions.TimeoutException:
        #    print("Unknown Error!!", end= " -- ")

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
    global aa_output_df
    aa_output_df = test_auto_data.loc[:, ["Sample Number", "AgreedValue"]]
    aa_output_df["AA_monthly_premium"] = ["-1"] * len(test_auto_data)
    aa_output_df["AA_yearly_premium"] = ["-1"] * len(test_auto_data)
    aa_output_df["AA_agreed_value_was_adjusted"] = [0] * len(test_auto_data)

    # save the number of cars in the dataset as a variable
    #num_cars = len(test_auto_data)
    num_cars = 2

    # loop through all cars in test spreadsheet
    for person_i in range(0, num_cars): 

        print(f"{person_i}: AA: ", end = "") # print out the iteration number

        # run on the ith car/person
        aa_auto_scrape(person_i)

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

    # scrape all of the insurance premiums for the given cars from aa
    auto_scape_all()

    # Close the browser window
    driver.quit()

main()
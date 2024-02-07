# webscraping related imports
import time
from selenium.common import exceptions
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys

# data management/manipulation related imports
import pandas as pd
from datetime import datetime
import math
import re
import os


## setting the working directory to be the folder this file is located in
# Get the absolute path of the current Python file
file_path = os.path.abspath(__file__)

# Get the directory of the current Python file
file_directory = os.path.dirname(file_path)

# Set the current working directory to be the directory of the Python file
os.chdir(file_directory)

## File path definitions
test_auto_data_xlsx = f"{file_directory}\\test_auto_data1.xlsx" # defines the path to the input dataset


"""
-------------------------
Useful functions
"""
def export_auto_dataset(num_datalines_to_export = len(test_auto_data_xlsx)):
    auto_dataset_for_export = test_auto_data.head(num_datalines_to_export) # get the given number of lines from the start
    auto_dataset_for_export.to_csv("scraped_auto_premium.csv", index=False)

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
            
            """
            print() # print just a newline character

            # printing a message to notify what is happened
            print(f"SELECTED: {final_car_variant.text}. ACTUAL: {data["Vehicle_year"]} {data["Manufacturer"]} {data["Model"]} {data["Model_type"]}" + 
                  f"{data["Model_series"]} with body type {data["Body_type"]} and {data["Num_speeds"]} {data["Automatic"]} and {data["Engine_size"]}L engine", end=" - ")
            """
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
                    return -1, -1

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
            Wait.until(EC.element_to_be_clickable( (By.XPATH, "//*[@id='quote']/fieldset[3]/div[1]/div[1]/ul/li[contains(text(), '{}')]".format(data["Suburb"])) )).click() # find the pop down option with the correct suburb

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
            return -1, -1 # we exit this option if the error page comes up (we cannot scrape for this person/car)
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
            monthly_premium, yearly_premium, adjusted_agreed_value = aa_auto_premium[0], aa_auto_premium[1], aa_auto_premium[2]
            print(monthly_premium, yearly_premium, end =" -- ")

    except:
        #try: # checks if the reason our code failed is because the 'we need more information' pop up appeareds
        Wait.until(EC.visibility_of_element_located( (By.XPATH, "//*[@id='ui-id-3' and text() = 'We need more information']") ) )
        print("Need more information", end= " -- ")
        #except exceptions.TimeoutException:
        #    print("Unknown Error!!", end= " -- ")

    end_time = time.time() # get time of end of each iteration
    print("Elapsed time:", round(end_time - start_time,2)) # print out the length of time taken

    # returning the adjusted agreed value
    try:
        return adjusted_agreed_value, monthly_premium, yearly_premium
    except UnboundLocalError: # if no value saved for adjusted_agreed value, then just return None
        return None




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
        
        # formatting the excess (rounded to the nearest option provided by AA)
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
            
            try: # try with the standard model type
                Wait.until(EC.element_to_be_clickable(( By.XPATH, "//div[@class='searchResultContainer']//*//label[contains(text(), ' {}')]".format(data["Model_type"]) ))).click() # wait until clickable, then click button to select final vehicle option
            except exceptions.TimeoutException: # if that doesn't work, try split into individual words
                model_type = data["Model_type"].split() # splits the model type string into words (allowing us to check for just 1st word)
                try: #try using just the first word of the model type
                    Wait.until(EC.element_to_be_clickable(( By.XPATH, "//div[@class='searchResultContainer']//*//label[contains(text(), ' {}')]".format(model_type[0])))).click()
                except: # try using just the last word of the model type
                    Wait.until(EC.element_to_be_clickable(( By.XPATH, "//div[@class='searchResultContainer']//*//label[contains(text(), ' {}')]".format(model_type[-1])))).click()
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
            print("Attempted to input agreed value larger than the maximum", end=" - ")
        elif data["Agreed_value"] < min_value:
            data["Agreed_value"] = min_value
            adjusted_agreed_value = min_value
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
        ami_auto_premium = ami_auto_scrape_premium(ami_auto_data_format(person_i))
        if ami_auto_premium != None: # if an actual result is returned
            monthly_premium, yearly_premium, adjusted_agreed_value = ami_auto_premium[0], ami_auto_premium[1], ami_auto_premium[2]
            print(monthly_premium, yearly_premium, end =" -- ")
    except:
        #try: # checks if the reason our code failed is because the 'we need more information' pop up appeareds
        Wait.until(EC.visibility_of_element_located( (By.XPATH, "//*[@id='ui-id-3' and text() = 'We need more information']") ) )
        print("Need more information", end= " -- ")
        #except exceptions.TimeoutException:
        #    print("Unknown Error!!", end= " -- ")

    end_time = time.time() # get time of end of each iteration
    print("Elapsed time:", round(end_time - start_time,2)) # print out the length of time taken

    # returning the adjusted agreed value
    try:
        return adjusted_agreed_value, monthly_premium, yearly_premium
    except UnboundLocalError: # if no value saved for adjusted_agreed value, then just return None
        return None




# defining a function that will scrape all of the ami cars
def state_auto_scrape_all():

    # defining a function which take the information from the spreadsheet and formats it so it can be used to scrape premium from ami website
    def state_auto_data_format(person_i):
        # splitting the street number into numeric and alphabeic aspects
        street_number = str(test_auto_data.loc[person_i,'Street_number'])
        street_number_numeric = remove_non_numeric(street_number)
        street_number_alphabetical = re.findall(r'\w', street_number)

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
            drivers_license_type = "NZ Full Licence" 
        elif test_auto_data.loc[person_i,'Licence'] == "RESTRICTED LICENCE":
            drivers_license_type = "NZ Restricted" 
        elif test_auto_data.loc[person_i,'Licence'] == "LEARNERS LICENCE":
            drivers_license_type = "NZ Learners" 
        else: # for for generic 'International' (non-NZ) licence
            drivers_license_type = "United Kingdom" 

        # formatting the number of years the person had had their drivers licence
        age_learners = test_auto_data.loc[person_i,'Age_learners_AMISTATE']

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
                    "Street_number":street_number,
                    "Street_number_numeric":street_number_numeric,
                    "Street_number_alphabetical":street_number_alphabetical,
                    "Street_name":street_name + " " + street_type,
                    "Suburb":suburb,
                    "Postcode":test_auto_data.loc[person_i,'Postcode'],
                    "Birthdate_day":int(birthdate.strftime("%d")),
                    "Birthdate_month":str(int(birthdate.strftime("%m"))),
                    "Birthdate_year":int(birthdate.strftime("%Y")),
                    "Sex":test_auto_data.loc[person_i,'Gender'],
                    "Drivers_license_type":drivers_license_type,
                    "Age_learners":age_learners, # years since driver got their learners licence
                    "Incidents_5_year":test_auto_data.loc[person_i,'Incidents_last5years_AMISTATE'],
                    "NZ_citizen_or_resident":test_auto_data.loc[person_i,'NZ_citizen_or_resident'],
                    "1_year_Visa":test_auto_data.loc[person_i,'Visa_at_least_1_year']
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
    def state_auto_scrape_premium(data):
        # internal function to select the correct car model variant
        def select_model_variant():
            # scraping these details from the webpage
            car_variant_options = tuple(driver.find_elements(By.XPATH, "//*[@id='vehicleFoundDetailsContainer']/div[2]/div"))

            if len(car_variant_options) == 1:
                return car_variant_options[0]
            else:
                print("Can't choose option!")
                time.sleep(10000)
            """
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

            # iterating through all other options to find one with least number of characters
            for car_variant in car_variant_options[1:]:
                if len(car_variant.text) < len(final_car_variant.text):
                    final_car_variant = car_variant
            
            # printing a message to notify what is happened
            print(f"SELECTED: {final_car_variant.text}. ACTUAL: {data["Vehicle_year"]} {data["Manufacturer"]} {data["Model"]} {data["Model_type"]}" + 
                  f"{data["Model_series"]} with body type {data["Body_type"]} and {data["Num_speeds"]} {data["Automatic"]} and {data["Engine_size"]}L engine", end=" - ")
            return final_car_variant
            """

        # internal function to select the correct address
        def select_correct_address():

            # inputs the address the car is kept at
            driver.find_element(By.ID, "situationOfRisk").send_keys(data["Street_number_numeric"] + " ") # input the street number
            driver.find_element(By.ID, "situationOfRisk").send_keys(data["Street_name"]) # input the street name


            # scraping all address options from webpage
            address_options = tuple(driver.find_elements(By.XPATH, "//ul[@id='ui-id-2']/li"))


            # define the address details list, in the order that we want to use them to filter out incorrect address options
            address_details_list = ["Suburb", "Postcode", "Street_number"]

            # iterate through all of the potential car specs that we can use to select the correct drop down option
            for address_detail in address_details_list:
                # initialise an empty list to store the selected car variants
                selected_addresses = [] 

                # check all addresses for this variable
                for address in address_options:

                    # Check that all of the known car details are correct (either starts with, ends with, or contains the details as a word in the middle of the text)
                    if data[address_detail].upper() in address.text.upper():
            
                        # if this car has correct details add it to the select list
                        selected_addresses.append(address) 


                # checking if we have managed to isolate one option
                if len(selected_addresses) == 1:
                    return selected_addresses[0]
                elif len(selected_addresses) > 1: # if there is still more than one option, remove all options that have already been filtered out
                    address_options = tuple(selected_addresses)

            print("Couldn't narrow down past:")
            for address in address_options:
                print(address.text)


        # Open the webpage
        #driver.get("https://secure.state.co.nz/car/?_gl=1*eiiger*_ga*MTc5MjY4MDcyMy4xNzA2NTgxMTI1*_ga_BLCG6HXQQC*MTcwNjU4MTEyNC4xLjEuMTcwNjU4MTg1OS41MS4wLjA.*_ga_F4QY8M53X3*MTcwNjU4MTEyNC4xLjEuMTcwNjU4MTg1OS4wLjAuMA..#/step1")

        driver.get("https://www.state.co.nz/car-insurance")
        Wait.until(EC.element_to_be_clickable((By.ID, "button-4715045f01")) ).click() # click button to get a quote (opens a new tab)
        driver.switch_to.window(driver.window_handles[-1])

        # attempt to input the car registration number (if it both provided and valid)
        registration_na = pd.isna(data["Registration_number"])
        if not registration_na: # if there is a registration number provided
            Wait10.until(EC.presence_of_element_located((By.ID, "regoInputBox")) ).send_keys(data["Registration_number"]) # input registration
            driver.find_element(By.ID, "searchByRegoBtn").click() # click 'Search' button

            # attempt to find the list element containing options for car variant (if present then we can continue)
            try: 

                Wait.until(EC.element_to_be_clickable( (By.XPATH,  "//*[@id='vehicleFoundDetailsContainer']/div[2]"))) # Options to select are below
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
                time.sleep(1) 
                driver.find_element(By.ID, "Year")
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
            
            # find the correct model variant
            model_variant_element = select_model_variant()

            # click the correct variant
            model_variant_element.click()
        
        except: # Selects the 1st option, if a Model type is not specified
            Wait.until(EC.element_to_be_clickable( (By.ID,  "searchedVehicleSpan_0"))).click() # wait until clickable, then click button to select final vehicle option
        
        time.sleep(3)

        # selects whether or not the car is used for business
        try:
            if data["Business_use"] == "No":
                driver.find_element(By.ID, "vehicleUsageNoLabel").click() # clicks "No" button
            else:
                driver.find_element(By.ID, "vehicleUsageYesLabel").click() # clicks "Yes" button
        except:
            print("Cannot click business use button", end=" -- ")
            return None

        time.sleep(1)

        # selects whether or not the car has an immobiliser
        try: # we 'try' this because the option to select Immobiliser only comes up on some cars (if there are some models of the car which don't)
            if data["Immobiliser"] == "Yes":
                driver.find_element(By.ID, "engineImmobiliserYesLabel").click() # clicks True button
            else:   
                driver.find_element(By.ID, "engineImmobiliserNoLabel").click() # clicks False button
        except:
            pass # if the button isn't present we move on
        
        time.sleep(1)
        
        # inputs and selects the correct address that the car is kept overnight at
        correct_address_element = select_correct_address()
        correct_address_element.click() # select the address

        time.sleep(2)

        # click 'Next' button
        driver.find_element(By.XPATH, '//*[@id="step1"]/div[3]/button').click()
        
        time.sleep(2) # wait for page to load

        # enter drivers birthdate
        driver.find_element(By.ID, "dobDayId_driver1").send_keys(data["Birthdate_day"]) # input day
        print(data["Birthdate_month"])
        Select(driver.find_element(By.ID, "dobMonthId_driver1")).select_by_value(data["Birthdate_month"]) # input month
        Wait.until(EC.element_to_be_clickable((By.ID, "dobYearId_driver1"))).send_keys(data["Birthdate_year"]) # input year

        time.sleep(1) # wait for page to load

        # select driver sex
        if data["Sex"] == "MALE":
            driver.find_element(By.ID, "genderMaleLabelId_driver1").click() # selects male
        else:
            driver.find_element(By.ID, "genderFemaleLabelId_driver1").click() # selects female

        # select 'No' for Do you have Contents, HomePack or Landlord insurance with State
        driver.find_element(By.ID, "mpdNoLabelId_driver1").click() # selects 'No

        # enter drivers licence info
        Select(driver.find_element(By.ID, "licenseId_driver1")).select_by_visible_text(data["Drivers_license_type"]) # select the correct drivers license type

        """
        if "International" in data["Drivers_license_type"]: # if international licence
            if data["NZ_citizen_or_resident"] == "Yes": # is the person NZ citizen/ perm resident
                driver.find_element(By.ID, 'prOrCitizen_1').click()
            else:
                driver.find_element(By.ID, 'notPrOrCitizen_1').click()
                if data["1_year_Visa"] == "Yes": # is the visa of the non-perm resident valid for more than one year
                    driver.find_element(By.ID, 'validVisa_1').click()
                else:
                    driver.find_element(By.ID, 'notValidVisa_1').click()
        """

        driver.find_element(By.ID, "licenseAgeId_driver1").send_keys(str(data["Age_learners"]))  # input age the person got their learners

        time.sleep(1) # wait for page to load

        # input if there have been any indicents
        if data["Incidents_5_year"] == "Yes":

            driver.find_element(By.ID, "incidentsYesLabelId_driver1").click() # opens incident type option box
            raise Exception("Incidents issue")
            """
            driver.find_element(By.XPATH, "//div[text()='{}']".format(data["Incident_type"])).click() # selects the driver incident type

            driver.find_element(By.ID, "DriverIncidentMonth_1").click() # opens incident month option box
            driver.find_element(By.XPATH, "//html//body//div[15]//div//div[text()='{}']".format(data["Incident_date_month"])).click() # selects the driver incident type
            driver.find_element(By.ID, "DriverIncidentYear_1").click() # opens incident year option box
            driver.find_element(By.XPATH, "//html//body//div[16]//div[text()='{}']".format(data["Incident_date_year"])).click() # selects the driver incident type
            """
        else:
            driver.find_element(By.ID, "incidentsNoLabelId_driver1").click() # clicks button saying that you have had no incidents


        time.sleep(5) # wait a bit for the page to load

        # click button to get quote 
        Wait.until(EC.element_to_be_clickable((By.ID, "save_quote"))).click() # wait until button clickable then click


        time.sleep(1000) # wait (for debugging)


        # wait until next page is loaded
        annual_risk_premium = Wait.until(EC.presence_of_element_located((By.ID, "annualRiskPremium")))

        # scrape the premium
        monthy_premium = float(driver.find_element(By.ID, "dollars").text.replace(",", "") + driver.find_element(By.ID, "cents").text)
        yearly_premium = float(annual_risk_premium.text.replace(",", "")[1:])

        # return the scraped premiums
        return monthy_premium, yearly_premium


    # loop through all cars in test spreadsheet  
    for person_i in range(3, len(test_auto_data)):
        start_time = time.time() # get time of start of each iteration

        print(person_i, ": ", end = "")
        # run on the ith car/person
        #try:
        ami_auto_premium = state_auto_scrape_premium(state_auto_data_format(person_i))
        if ami_auto_premium != None: # if an actual result is returned
            monthly_premium, yearly_premium = ami_auto_premium[0], ami_auto_premium[1]
            print(monthly_premium, yearly_premium, end =" -- ")
        """
        except:
            try: # checks if the reason our code failed is because the 'we need more information' pop up appeareds
                Wait.until(EC.visibility_of_element_located( (By.XPATH, "//*[@id='ui-id-3' and text() = 'We need more information']") ) )
                print("Need more information", end= " -- ")
            except exceptions.TimeoutException:
                print("Unknown Error!!", end= " -- ")
        """

        end_time = time.time() # get time of end of each iteration
        print("Elapsed time:", round(end_time - start_time,2)) # print out the length of time taken

        # refresh the page every 100 iterations, to (hopefully) prevent memory overloads
        if person_i % 100 == 0 and person_i > 0:
            driver.refresh()
        
        try:
            # delete all cookies to reset the page
            driver.delete_all_cookies()
        except exceptions.TimeoutException: # if we timeout while trying to reset the cookies

            print("New Webdriver window", end=" -- ")
            driver.quit() # quit this current driver
            load_webdriver() # open a new webdriver session




# defining a function that will scrape all of the tower cars
def tower_auto_scrape(person_i):
# defining a function which take the information from the spreadsheet and formats it so it can be used to scrape premium from tower website
    def tower_auto_data_format(person_i):
        # saving manufacturer as a variable
        manufacturer = str(test_auto_data.loc[person_i,'Manufacturer']).title()

        # formatting model type
        model_type = test_auto_data.loc[person_i,'Type']
        if pd.isna(model_type):
            model_type = ""
        elif "XL Dual Cab" in model_type:
            model_type = model_type.replace("Dual", "Double")

        # getting the street address
        street_name = test_auto_data.loc[person_i,'Street_name']
        street_type = test_auto_data.loc[person_i,'Street_type']
        suburb = test_auto_data.loc[person_i,'Suburb']
        if "(" in street_name:
            street_name = street_name.split("(")[0].strip()
        if "MT " in suburb:
            suburb = suburb.replace("MT", "MOUNT")

        # formatting unit number
        unit_number = test_auto_data.loc[person_i,'Unit_number']
        if pd.isna(unit_number): # if there is no unit number (is na), set the unit number varibale to an empty string
            unit_number = ""

        # formatting engine size
        engine_size = f"{round(test_auto_data.loc[person_i,'CC']/1000, 1)}"
        if engine_size == "0.0": # if there is no cubic centimetres engine measurement for car (is electric)
            engine_size = ""

        # formatting car model type (for when the label is just C)
        model = test_auto_data.loc[person_i,'Model']
        if model == "C":
            model += str(engine_size) # add on the number of '10 times litres' in the engine
        
        ## formatting gearbox info
        automatic = test_auto_data.loc[person_i,'Gearbox']

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
        birthdate = test_auto_data.loc[person_i,'DOB']

        # formatting the excess (rounded to the nearest option provided by AA)
        excess = float(test_auto_data.loc[person_i,'Excess']) # convert into a floating point value (if it is not already one)
        excess_options = [400, 500, 750, 1000] # defines a list of the acceptable 
        # choose the largest excess option for which the customers desired excess is still larger
        excess_index = 0
        while excess > excess_options[excess_index] and excess_index < 3: # 3 is the index of the largest option, so should not iterate up further if the index has value 3
            excess_index += 1

        # formatting the policy start date
        policy_start_date = datetime.strftime(test_auto_data.loc[person_i, "PolicyStartDate"], "%d/%m/%Y") # ensure that the date is in the correct format 'DD/MM/YYYY'

        # define a dict to store information for a given person and car for ami
        tower_data = {"Registration_number":test_auto_data.loc[person_i,'Registration'],
                    "Manufacturer":manufacturer,
                    "Model":str(model).title(),
                    "Model_type":model_type,
                    "Vehicle_year":test_auto_data.loc[person_i,'Vehicle_year'],
                    "Body_type":str(test_auto_data.loc[person_i,'Body']).title(),
                    "Engine_size":engine_size,
                    "Num_speeds":num_speeds,
                    "Automatic":automatic,
                    "Business_use":test_auto_data.loc[person_i,'BusinessUser'],
                    "Unit":unit_number,
                    "Street_number":test_auto_data.loc[person_i,'Street_number'],
                    "Street_name":f"{street_name} {street_type}".title(),
                    "Suburb":test_auto_data.loc[person_i,'Suburb'].strip().title(),
                    "Postcode":test_auto_data.loc[person_i,'Postcode'],
                    "Birthdate_day":int(birthdate.strftime("%d")),
                    "Birthdate_month":birthdate.strftime("%m"),
                    "Birthdate_year":int(birthdate.strftime("%Y")),
                    "Sex":test_auto_data.loc[person_i,'Gender'],
                    "Incidents_3_year":test_auto_data.loc[person_i,'Incidents_last3years_AA'],
                    "Exclude_under_25":test_auto_data.loc[person_i,'ExcludeUnder25'],
                    "Cover_type":test_auto_data.loc[person_i,'CoverType'],
                    "Agreed_value":int(round(test_auto_data.loc[person_i,'AgreedValue'])), # rounds the value to nearest whole number then converts to an integer
                    "Excess_index":excess_index,
                    "Modifications":test_auto_data.loc[person_i,'Modifications'],
                    "Immobiliser":test_auto_data.loc[person_i,'Immobiliser_alarm'], 
                    "Finance_purchase":test_auto_data.loc[person_i,'FinancePurchase'],
                    "Policy_start_date":str(policy_start_date)
                    }
        
        # adding info on the date and type of incident to the ami_data dictionary ONLY if the person has had an incident within the last 5 years
        incident_date = test_auto_data.loc[person_i,'Date_of_incident']
        if tower_data["Incidents_3_year"] == "Yes":
            tower_data["Incident_year"] = int(incident_date.strftime("%Y"))

            incident_type = test_auto_data.loc[person_i,'Type_incident'].lower() # initialising the variable incident type

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
        # Open the webpage
        driver.get("https://my.tower.co.nz/quote/car/page1")

        # wait for page to load
        time.sleep(1)

        # selects whether or not the car is used for business
        try: # if error go to except
            if data["Business_use"] == "Yes":
                Wait10.until(EC.element_to_be_clickable( (By.ID, "btnvehicleUsedForBusiness-0") )).click() # clicks "Yes" business use button
            elif data["Business_use"] == "No":
                Wait10.until(EC.element_to_be_clickable( (By.ID, "btnvehicleUsedForBusiness-1") )).click() # clicks "No" business use button
            else:
                Wait10.until(EC.element_to_be_clickable( (By.ID, "btnvehicleUsedForBusiness-2") )).click() # clicks "Sometimes" business use button
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
                        print("Car found by registration number is incorrect!", end=" -- ")
                        raise ValueError("Car found by registration number is incorrect!")
                    
                    # if the car model type, transmission type, number of speeds and engine size (num of litres) are not correct, then enter details manually
                    extra_registration_found_car_text = registration_found_car.find_element(By.XPATH, ".//p").text
                    if not ( (str(data["Model_type"]) in extra_registration_found_car_text) and (str(data["Automatic"]) in extra_registration_found_car_text) and (str(data["Num_speeds"]) in extra_registration_found_car_text) and (str(data["Engine_size"]) in extra_registration_found_car_text) ):
                        print("Car found by registration number has several options", end=" -- ")
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

                    time.sleep(2) # wait for page to load

                    # click the button to select the car manufacturer in the dropdown (i just click the 1st drop down option because I assume this must be the correct one)
                    Wait.until(EC.presence_of_element_located( (By.XPATH, "//*[@id='carMakes-menu-list']/li/div/div[2]/div") )).click() 
            except exceptions.TimeoutException:
                print(f"CANNOT FIND {data["Manufacturer"]}", end=" -- ")
                return None # return None if can't scrape

            # inputting car model
            try:
                # wait until car model input box is clickable, then input the car model
                car_model_text_input = Wait.until(EC.presence_of_element_located((By.ID, "carModels")))
                if car_model_text_input.get_attribute("value") == "": # checks the input fields value is currently empty
                    car_model_text_input.send_keys(data["Model"]) 

                    time.sleep(2) # wait for page to load
                    
                    # wait until button which has the correct car model information is clickable, then click (i just click the 1st drop down option because I assume this must be the correct)
                    Wait10.until(EC.element_to_be_clickable( (By.XPATH, "//*[@id='carModels-menu-list']/li[1]/div/div[2]/div") )).click() 
            except exceptions.TimeoutException:
                print(f"CANNOT FIND {data["Manufacturer"]} MODEL {data["Model"]}", end=" -- ")
                return None # return None if can't scrape

            # inputting car year
            try:
                car_model_text_input = Wait.until(EC.presence_of_element_located( (By.ID, "carYears") )) # find car year input box
                if car_model_text_input.get_attribute("value") == "": # checks the input fields value is currently empty
                    car_model_text_input.send_keys(str(data["Vehicle_year"])) # inputs the year 

                    time.sleep(1) # wait for page to load

                    Wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='carYears-menu-list']/li[1]"))).click() # clicking the button which has the correct car year information
            except exceptions.TimeoutException:
                print(f"CANNOT FIND {data["Manufacturer"]} {data["Model"]} FROM YEAR {data["Vehicle_year"]}", end=" -- ")
                return None # return None if can't scrape


            # inputting car body style
            try:
                body_type_text_input = Wait.until(EC.presence_of_element_located((By.ID, "carBodyStyles"))) # find the car body type input box and
                if body_type_text_input.get_attribute("value") == "": # checks the input fields value is currently empty

                    body_type_text_input.send_keys(data["Body_type"]) # inputs the body type

                    time.sleep(1) # wait for page to load

                    Wait.until(EC.element_to_be_clickable( (By.XPATH, "//*[@id='carBodyStyles-menu-list']/li[1]/div/div[2]/div") )).click() # clicking the button which has the correct car body style information
            except exceptions.TimeoutException: # if code timeout while waiting for element
                print(f"CANNOT FIND {data["Vehicle_year"]} {data["Manufacturer"]} {data["Model"]} WITH BODY TYPE {data["Body_type"]}", end=" -- ")
                return None # return None if can't scrape

            # inputting car vehicle type
            try:
                car_model_type_text_input = Wait.until(EC.presence_of_element_located((By.ID, "carVehicleTypes"))) # find the model type input box and then input the model type
                if car_model_type_text_input.get_attribute("value") == "": # checks the input fields value is currently empty
                    car_model_type_text_input.send_keys(data["Model_type"]) # inputs the body type

                    time.sleep(1) # wait for page to load

                    # clicking the button which has the correct model type information, as well as the correct Gearbox (Automatic and Num_speeds) as well as Engine_Size
                    Wait.until(EC.element_to_be_clickable( (By.XPATH, f"//*[@id='carVehicleTypes-menu-list']/li/div[contains(@value, '{data["Body_type"]}') and " + 
                                                            f"contains(@value, '{data["Automatic"]}') and " + 
                                                            f"contains(@value, '{data["Num_speeds"]}') and " +
                                                            f"contains(@value, '{data["Engine_size"]}')]") )).click()       
            except exceptions.TimeoutException:
                
                # try searching in a less strict way
                try:
                    if data["Automatic"] == "Auto" or data["Automatic"] == "CVT": # some cars that are autos are labeled as CVT's
                        Wait.until(EC.element_to_be_clickable( (By.XPATH, f"//*[@id='carVehicleTypes-menu-list']/li/div[contains(@value, '{data["Body_type"]}') and " + 
                                                            f"( contains(@value, 'Auto') or contains(@value, 'CVT') or contains(@value, 'Skyactiv-Drive')) and " + # note Skyactive-Drive is a Mazda specific automatic
                                                            f"contains(@value, '{data["Engine_size"]}')]") )).click()
                    elif data["Automatic"] == "Man": # if the car is a manual
                        Wait.until(EC.element_to_be_clickable( (By.XPATH, f"//*[@id='carVehicleTypes-menu-list']/li/div[contains(@value, '{data["Body_type"]}') and " + 
                                                            f"( contains(@value, 'Man') or contains(@value, 'DCT') or contains(@value, 'Skyactiv-MT')) and " + # note Skyactive-MT is a Mazda specific manual
                                                            f"contains(@value, '{data["Engine_size"]}')]") )).click()
                except exceptions.TimeoutException:
                    print(f"CANNOT FIND {data["Vehicle_year"]} {data["Manufacturer"]} {data["Model"]} {data["Model_type"]} {data["Body_type"]}, {data["Num_speeds"]} {data["Automatic"]} WITH ENGINE SIZE {data["Engine_size"]}", end=" -- ")
                    return None # return None if can't scrape
        
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
            return -1, -1

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
        
        # check if our attempted agreed value is valid. if not, round up/down to the min/max value
        if data["Agreed_value"] > max_value:
            data["Agreed_value"] = max_value
            adjusted_agreed_value = max_value # saves the adjusted agreed value to return later
            print("Attempted to input agreed value larger than the maximum", end=" - ")
        elif data["Agreed_value"] < min_value:
            data["Agreed_value"] = min_value
            adjusted_agreed_value = min_value # saves the adjusted agreed value to return later
            print("Attempted to input agreed value smaller than the minimum", end=" - ")


        # inputs the agreed value input the input field (after making sure its valid)
        agreed_value_input = driver.find_element(By.ID, "agreedValueNewSliderField") # find the input field for the agreed value
        agreed_value_input.send_keys(Keys.CONTROL, "a") # select all current value
        agreed_value_input.send_keys(data["Agreed_value"]) # input the desired value, writing over the (selected) current value
        driver.find_element(By.ID, "agreedValueNewSliderBtn").click() # click the 'Update agreed value' button


        time.sleep(1) # wait for page to load


        # input the persons desired level of excess
        Wait.until(EC.element_to_be_clickable( (By.ID, f"btnexcess-{data["Excess_index"]}"))).click()

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
        next_button.click() # click 'Next: Summary' button

        time.sleep(4) # wait for the page to load

        # move onto the next page "People"
        Wait.until(EC.element_to_be_clickable((By.ID, "btnSubmitPage"))).click()

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
        try:
            return round(monthly_premium, 2), round(yearly_premium, 2), adjusted_agreed_value
        except UnboundLocalError: # if no value saved for adjusted_agreed value, then just return None for it
            return round(monthly_premium, 2), round(yearly_premium, 2), None


    start_time = time.time() # get time of start of each iteration

    # run on the ith car/person
    try:
        auto_premiums = tower_auto_scrape_premium(tower_auto_data_format(person_i))
        if auto_premiums != None: # if an actual result is returned
            monthly_premium, yearly_premium, adjusted_agreed_value = auto_premiums[0], auto_premiums[1], auto_premiums[2]
            print(monthly_premium, yearly_premium, end =" -- ")
    except:
        #try: # checks if the reason our code failed is because the 'we need more information' pop up appeareds
        Wait.until(EC.visibility_of_element_located( (By.XPATH, "//*[@id='ui-id-3' and text() = 'We need more information']") ) )
        print("Need more information", end= " -- ")
        #except exceptions.TimeoutException:
            #print("Unknown Error!!", end= " -- ")
        #    raise Exception("Unknown Error")


    end_time = time.time() # get time of end of each iteration
    print("Elapsed time:", round(end_time - start_time,2)) # print out the length of time taken

        # returning the adjusted agreed value
    try:
        return adjusted_agreed_value, monthly_premium, yearly_premium
    except UnboundLocalError: # if no value saved for adjusted_agreed value, then just return None
        return None
    




def auto_scape_all():
    # run the scraper for person_i on the given company website
    def run_company_scrapers(person_i, company):
        
        print(f"{person_i}: {company}: ", end = "")

        # run on the ith car/person for the given company
        if company == "AA":
            adjusted_agreed_value, monthly_premium, yearly_premium = aa_auto_scrape(person_i)
        elif company == "AMI":
            adjusted_agreed_value, monthly_premium, yearly_premium = ami_auto_scrape(person_i)
        elif company == "Tower":
            adjusted_agreed_value, monthly_premium, yearly_premium = tower_auto_scrape(person_i)
        
        # save the scraped monthly and yearly premiums to the test_auto_data pandas dataframe
        test_auto_data.loc[person_i, f"{company}_monthly_premium"], test_auto_data.loc[person_i, f"{company}_yearly_premium"] = monthly_premium, yearly_premium
        
        
        return adjusted_agreed_value
    
    # save the adjusted agreed value in the spreadsheet (so that all companies are using the same agreed value)
    def save_adjusted_agreed_value(insurance_companies, current_company, adjusted_agreed_value, person_i):
        # save the adjusted agreed value
        global test_auto_data
        test_auto_data.loc[person_i,'AgreedValue'] = adjusted_agreed_value

        # iterate through all previous companies using this new adjusted agreed value for person_i
        for company in insurance_companies:

            # stop once we have iterated through all previous companies
            if company == current_company:
                return
            else:
                adjusted_agreed_value = run_company_scrapers(person_i, company)

                if adjusted_agreed_value != None:
                    save_adjusted_agreed_value(insurance_companies, company, adjusted_agreed_value, person_i)

    # reformats the dataset to add columns to store the scraped insurance premiums
    def reformat_test_auto_data(insurance_companies, num_cars = len(test_auto_data)):
        # add a new column 'Company_monthly_premium' and Company_yearly_premium' for each company in insurance companies
        for company in insurance_companies:
            global test_auto_data
            test_auto_data[f"{company}_monthly_premium"] = [None] * num_cars
            test_auto_data[f"{company}_yearly_premium"] = [None] * num_cars

    # save the number of cars in the dataset as a variable
    #num_cars = len(test_auto_data)
    num_cars = 1

    # define a list of insurance companies to iterate through
    insurance_companies = ["AA", "AMI", "Tower"]
    
    reformat_test_auto_data(insurance_companies)

    # estimate the number of seconds testing all cars on each company website will take
    approximate_total_times = [(time * num_cars) for time in [50, 40, 65]]
    total_time_hours = sum(approximate_total_times) / 3600 # convert seconds to hours
    total_time_minutes = round((total_time_hours - int(total_time_hours)) * 60)
    total_time_hours = math.floor(total_time_hours)

    print(f"Program will take approximately {total_time_hours} hours and {total_time_minutes} minutes to scrape the premiums for {num_cars} cars for AA, AMI and Tower", end="\n\n\n")

    # loop through all cars in test spreadsheet
    for person_i in range(0, num_cars): 

        # iterate through all the different insurance providers
        for company in insurance_companies:

            # run the scraper for person_i on the given company website
            adjusted_agreed_value = run_company_scrapers(person_i, company)

            # if the agreed value was adjusted to work with this company (if the agreed value was larger/ smaller than the upper/ lower accepted limits)
            if adjusted_agreed_value != None: 
                save_adjusted_agreed_value(insurance_companies, company, adjusted_agreed_value, person_i) # save the adjusted agreed value in the spreadsheet (so that all companies are using the same agreed value)

            # delete all cookies to reset the page
            try:
                driver.delete_all_cookies()
            except exceptions.TimeoutException: # if we timeout while trying to reset the cookies

                print("\n\nNew Webdriver window\n")
                driver.quit() # quit this current driver
                load_webdriver() # open a new webdriver session
    
    export_auto_dataset(num_cars)


def main():
    # reads in the test data for car insurance inputs
    global test_auto_data
    test_auto_data = pd.read_excel(test_auto_data_xlsx, dtype={"Postcode":"int"})

    # pads out the front of postcodes with zeroes (as excel removes leading zeros)
    test_auto_data['Postcode'] = test_auto_data['Postcode'].apply(postcode_reformat) 

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
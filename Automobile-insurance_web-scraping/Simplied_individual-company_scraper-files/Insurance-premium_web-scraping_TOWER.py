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

# importing for more natural string comparisons
from fuzzywuzzy import fuzz

# importing several general functions (which are defined in the seperate python file called funct_defs)
import funct_defs

"""
-------------------------
Useful functions
"""

# a function to open the webdriver (chrome simulation)
def load_webdriver():

    # defines the webdriver variables
    global driver, Wait3, Wait10
    driver, Wait3, Wait10 = funct_defs.load_webdriver()


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

        if "XL Dual Cab" in model_type:
            model_type = model_type.replace("Dual", "Double")

        # formatting model series
        model_series = str(test_auto_data_df.loc[person_i,'Series'])

        # getting the street address variables
        street_name = test_auto_data_df.loc[person_i,'Street_name']
        street_type = test_auto_data_df.loc[person_i,'Street_type']
        suburb = test_auto_data_df.loc[person_i,'Suburb']


        if "(" in street_name:
            street_name = street_name.split("(")[0].strip()
        if "MT " in suburb:
            suburb = suburb.replace("MT", "MOUNT")

        # formatting unit number
        unit_number = test_auto_data_df.loc[person_i,'Unit_number']

        # formatting engine size
        engine_size = f"{round(float(test_auto_data_df.loc[person_i,'CC'])/1000, 1)}"
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
        elif "Sports" in automatic: # all the different types of automatic transmissions (from the test excel spreadsheeet) and how they are labeled
            automatic = "Spts Auto" 
        elif "Automatic" in automatic or "DSG" in automatic: # all the different types of automatic transmissions (from the test excel spreadsheeet) and how they are labeled
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
                    "Suburb":suburb.strip().title(),
                    "Postcode":test_auto_data_df.loc[person_i,'Postcode'],
                    "Birthdate_day":int(birthdate.strftime("%d")),
                    "Birthdate_month":birthdate.strftime("%m"),
                    "Birthdate_year":int(birthdate.strftime("%Y")),
                    "Sex":test_auto_data_df.loc[person_i,'Gender'],
                    "Incidents_3_year":int(test_auto_data_df.loc[person_i,'Incidents_last3years_AA']),
                    "Exclude_under_25":test_auto_data_df.loc[person_i,'ExcludeUnder25'],
                    "Cover_type":test_auto_data_df.loc[person_i,'CoverType'],
                    "Agreed_value":int(round(float(test_auto_data_df.loc[person_i,'AgreedValue']))), # rounds the value to nearest whole number then converts to an integer
                    "Excess_index":excess_index,
                    "Modifications":test_auto_data_df.loc[person_i,'Modifications'],
                    "Immobiliser":test_auto_data_df.loc[person_i,'Immobiliser_alarm'], 
                    "Finance_purchase":test_auto_data_df.loc[person_i,'FinancePurchase'],
                    "Policy_start_date":datetime.today().strftime("%d/%m/%Y"),
                    "Email":str(test_auto_data_df.loc[person_i,'Email']),
                    "Phone_number":str(test_auto_data_df.loc[person_i,'Phone Number']),
                    "Business_or_trust":str(test_auto_data_df.loc[person_i,'Policy Owned by a Business or Trust']),
                    "License_suspended":str(test_auto_data_df.loc[person_i,'License Suspended in Last 7 Years']),
                    "Insurance_refused_7_years":str(test_auto_data_df.loc[person_i,'Insurance Refused In Last 7 Years']).upper() == "YES",
                    "Claim_refused_7_years":str(test_auto_data_df.loc[person_i,'Claim Refused In Last 7 Years']).upper() == "YES",
                    "Crime_7_years":str(test_auto_data_df.loc[person_i,'Crime in Last 7 Years']).upper() == "YES",
                    "Additional_drivers":test_auto_data_df.loc[person_i, "Additional Drivers"].upper() == "YES"
                    }
        
        # formatting the bank that finance was got from
        if tower_data["Finance_purchase"].upper() == "YES":
            tower_data["Bank"] = str(test_auto_data_df.loc[person_i,'Finance Bank'])

        # formatting the firstname
        if pd.isna(test_auto_data_df.loc[person_i,'First_name']) or test_auto_data_df.loc[person_i,'First_name'] == "":
            if tower_data["Sex"] == "MALE":
                tower_data["First_name"] = "John"
            else:
                tower_data["First_name"] = "Jane"
        else:
            tower_data["First_name"] = str(test_auto_data_df.loc[person_i,'First_name'])
        
        # formatting the surname
        if pd.isna(test_auto_data_df.loc[person_i,'Surname']) or test_auto_data_df.loc[person_i,'Surname'] == "":
            tower_data["Surname"] = "Doe"
        else:
            tower_data["Surname"] = str(test_auto_data_df.loc[person_i,'Surname'])

        # adding info on the date and type of incident to the tower_data dictionary ONLY if the person has had an incident within the last 5 years
        if tower_data["Incidents_3_year"] > 0:

            # saving all the incident dates
            for i in range(1, tower_data["Incidents_3_year"] + 1):
                tower_data[f"Incident{i}_year"] = int(test_auto_data_df.loc[person_i,f'Date_of_incident{i}'].strftime("%Y"))

            # saving the incident type
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
        # function to construct the string that summarises the details about a car
        def db_car_details_string_constructor(data):
            details_list = ["Model_type", "Body_type","Automatic", "Num_speeds", "Engine_size"]

            # initialising the output string variable 
            output_string = f"{data["Model_series"].strip()}"

            # adding a space to the variable only if Model_series is not an empty string
            if data["Model_series"] != "":
                output_string += " "
            
            for index in range(len(details_list)):
                output_string += f"{data[details_list[index]].strip()}"
                    
                # adding a space to the variable only if the given detail is not an empty string and is not the last detail
                if index < len(details_list) - 1 and data[details_list[index]].strip() != "":
                    output_string += " "
            
            return output_string

        # attempt to click business use button
        def handle_business_use_button():
            # Checks if the "More info required" popup is present
            try:
                Wait3.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="recaptchaDialog"]/section/div/h5[text()="More info required!"]')))
            except exceptions.TimeoutException:
                pass
            else:
                driver.refresh() # refresh the page

            # selects whether or not the car is used for business
            if data["Business_use"] == "Yes":
                Wait10.until(EC.element_to_be_clickable( (By.ID, "btnvehicleUsedForBusiness-0") )).click() # clicks "Yes" business use button
            elif data["Business_use"] == "No":
                Wait10.until(EC.element_to_be_clickable( (By.ID, "btnvehicleUsedForBusiness-1") )).click() # clicks "No" business use button
            else:
                Wait10.until(EC.element_to_be_clickable( (By.ID, "btnvehicleUsedForBusiness-2") )).click() # clicks "Sometimes" business use button

        def select_model_variant(db_car_details = db_car_details_string_constructor(data), xpath = '//*[@id="carVehicleTypes-menu-list"]/li'):
            # scraping these details from the webpage
            car_variant_options = tuple(driver.find_elements(By.XPATH, xpath))

            # filter out all options where the engine size is incorrect
            car_variant_options = [option for option in car_variant_options if (data["Engine_size"] in option.text)]

            # if there are no car variant options with the correct number of speeds and correct engine size
            if len(car_variant_options) == 0:
                raise Exception("Unable to find car variant")

            time.sleep(0.5) # waiut for page to load

            # get a list of the similarity scores of our car variant option, compared with the string summarising the info from the database
            car_variant_accuracy_score = [fuzz.ratio(db_car_details, option.text) for option in car_variant_options]

            # save the highest accuarcy score
            max_value = max(car_variant_accuracy_score)

            # get the car variant option(s) that match the data the best
            car_variant_options = [car_variant_options[index] for index, score in enumerate(car_variant_accuracy_score) if score == max_value]

            if len(car_variant_options) > 1:
                print("Unable to fully narrow down", end=" - ")
                tower_output_df.loc[person_i, "TOWER_Error_code"] = "Several Car Variant Options Warning"
            
            # saving the select model variant to the output df
            tower_output_df.loc[person_i, "TOWER_selected_car_variant"] = f"{data["Manufacturer"]} {data["Model"]} {data["Vehicle_year"]} {car_variant_options[0].text}"

            # return the (1st) best matching car variant option
            return car_variant_options[0]

        def enter_registration_number():
            if data["Registration_number"] == "": # if the vehicle registration number is NA then raise an exception (go to except block, skip rest of try)
                raise ValueError("Registration_NA")
            else:
                # attempt to input the license plate number. If it doesn't work then raise value error to go enter the car details manually
                driver.find_element(By.ID, "txtLicencePlate").send_keys(data["Registration_number"]) # input registration number
                driver.find_element(By.ID, "btnSubmitLicencePlate").click() # click seach button
                

                time.sleep(2) # wait for page to load
                
                # checking if the 'More info required' info box appears (if if does we return the Website Does Not Quote For This Car Variant error code)
                try:
                    Wait3.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="carLookupError"]/div/p[text()="More info required"]')))
                    raise Exception("Website Does Not Quote For This Car Variant/ Person")
                
                except exceptions.TimeoutException:
                    pass

                try:
                    Wait3.until_not(EC.presence_of_element_located( (By.XPATH, "//*[@id='carLookupError']/div/p[text() = 'No record found']") )) # checks that the carLookupError pop up does NOT appear
                except exceptions.TimeoutException: # if the carLookupError pop up appears
                    raise ValueError("Registration_Invalid")


                # try to select the correct model variant from a pop down list of options
                try:
                    # checking if an options box to select has appeared
                    Wait3.until(EC.visibility_of_element_located((By.XPATH,'//*[@id="questionCarLookup"]/div[4]/div[2]/fieldset/div[1]/label/div')))

                    # finding the correct model variant (on the assumption that the list of options has appeared)
                    model_variant_element = select_model_variant(db_car_details=f"{data["Manufacturer"]} {data["Model"]} {data["Vehicle_year"]} {data["Model_type"]} {data["Model_series"]} {data["Body_type"]} {data["Automatic"]} {data["Num_speeds"]} {data["Engine_size"]}",
                                                                  xpath = '//*[@id="questionCarLookup"]/div[4]/div[2]/fieldset/div')
                    
                    if model_variant_element != None and model_variant_element != "Unable to find car variant": # if a well matching car variant was found
                        model_variant_element.click()
                    else: # if there are no options that match with the provided data
                        raise ValueError("Car found by registration number is incorrect!")
                
                except exceptions.TimeoutException: # if the list of options is not present, check if the correct car model has been already selected
                    # find the elements which lists the cars details, after having input the registration number
                    registration_found_car = Wait10.until(EC.presence_of_element_located((By.XPATH, "//*[@id='questionCarLookup']/div[4]/div[@class='car-results']")) )
                    
                    # constructing a string to represent the car details from the input data
                    db_car_details = f"{data["Manufacturer"]} {data["Model"]} {data["Body_type"]} {data["Vehicle_year"]} {data["Model_type"]} {data["Model_series"]} {data["Automatic"]} {data["Num_speeds"]} {data["Engine_size"]}"

                    # check the similarity between the details on the website and the actual car details in the database, if similarly is not over 80% then enter details manually
                    if fuzz.token_set_ratio(db_car_details, registration_found_car.text) < 80:
                        raise ValueError("Car found by registration number is incorrect!")
                    
                    # saving the select model variant to the output df
                    tower_output_df.loc[person_i, "TOWER_selected_car_variant"] = registration_found_car.text

        def enter_car_details_manually():
            
            # defining a function to check if our model selection is valid

            time.sleep(1) # wait for page to load
            
            # Find the button "Enter your car's details" and click it (only if not already present)
            try:
                # find car manufacturer input box and input the company that manufactures the car
                manufacturer_text_input = driver.find_element(By.ID, "carMakes")
            except:
                Wait3.until(EC.element_to_be_clickable( (By.ID, "lnkEnterMakeModel") )).click()

                # find car manufacturer input box and input the company that manufactures the car
                manufacturer_text_input = driver.find_element(By.ID, "carMakes")

            # inputting the car manufacturer
            try:

                if manufacturer_text_input.get_attribute("value") == "": # checks the input fields value is currently empty
                    manufacturer_text_input.send_keys(data["Manufacturer"])

                    # wait until the options are ready to be clicked
                    Wait10.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="carMakes-wrapper"]/div/i'))) # wait for page to load
                    time.sleep(1)
                    
                    # click the button to select the car manufacturer in the dropdown (i just click the 1st drop down option because I assume this must be the correct one)
                    Wait3.until(EC.element_to_be_clickable( (By.XPATH, "//*[@id='carMakes-menu-list']/li/div/div[2]/div") )).click() 
            except exceptions.TimeoutException:
                raise Exception(f"Unable to find car variant: MANUFACTURER {data["Manufacturer"]}")

            # inputting car model
            try:
                # wait until car model input box is clickable, then input the car model
                car_model_text_input = Wait3.until(EC.presence_of_element_located((By.ID, "carModels")))
                if car_model_text_input.get_attribute("value") == "": # checks the input fields value is currently empty
                    car_model_text_input.send_keys(data["Model"]) 
                    
                    # wait until the options are ready to be clicked
                    Wait10.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="carModels-wrapper"]/div/i'))) # wait for page to load
                    time.sleep(1)

                    # wait until button which has the correct car model information is clickable, then click (i just click the 1st drop down option because I assume this must be the correct)
                    Wait3.until(EC.element_to_be_clickable( (By.XPATH, "//*[@id='carModels-menu-list']/li[1]/div/div[2]/div") )).click() 
            except exceptions.TimeoutException:
                raise Exception(f"Unable to find car variant: {data["Manufacturer"]} MODEL {test_auto_data_df.loc[person_i,'Model']}")

            # inputting car year
            try:
                car_model_text_input = Wait3.until(EC.presence_of_element_located( (By.ID, "carYears") )) # find car year input box
                if car_model_text_input.get_attribute("value") == "": # checks the input fields value is currently empty
                    car_model_text_input.send_keys(str(data["Vehicle_year"])) # inputs the year 

                    # wait until the options are ready to be clicked
                    Wait3.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="carYears-wrapper"]/div/i'))) # wait for page to load
                    time.sleep(1)

                    Wait3.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='carYears-menu-list']/li[1]"))).click() # clicking the button which has the correct car year information
            except exceptions.TimeoutException:
                raise Exception(f"Unable to find car variant: {data["Manufacturer"]} {test_auto_data_df.loc[person_i,'Model']} FROM YEAR {data["Vehicle_year"]}")


            # inputting car body style
            try:
                body_type_text_input = Wait3.until(EC.presence_of_element_located((By.ID, "carBodyStyles"))) # find the car body type input box and
                if body_type_text_input.get_attribute("value") == "": # checks the input fields value is currently empty

                    body_type_text_input.send_keys(data["Body_type"]) # inputs the body type

                    # wait until the options are ready to be clicked
                    Wait3.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="carBodyStyles-wrapper"]/div/i'))) # wait for page to load
                    time.sleep(1)

                    Wait3.until(EC.element_to_be_clickable( (By.XPATH, "//*[@id='carBodyStyles-menu-list']/li[1]/div/div[2]/div") )).click() # clicking the button which has the correct car body style information
            except exceptions.TimeoutException: # if code timeout while waiting for element
                raise Exception(f"Unable to find car variant: {data["Vehicle_year"]} {data["Manufacturer"]} {test_auto_data_df.loc[person_i,'Model']} WITH BODY TYPE {data["Body_type"]}")

            # inputting car vehicle type
            car_model_type_text_input = Wait3.until(EC.presence_of_element_located((By.ID, "carVehicleTypes"))) # find the model type input box and then input the model type

            if car_model_type_text_input.get_attribute("value") == "": # checks the input fields value is currently empty

                # if the body type is not empty
                if data["Model_type"] != "":
                        car_model_type_text_input.send_keys(data["Model_type"]) # inputs the model type
                else: # if the body type is not empty, then open the drop down by just clicking the element
                    car_model_type_text_input.click()
                    time.sleep(0.5) # wait for page to load
                

                # checking if an options box to select has appeared
                Wait3.until(EC.visibility_of_element_located((By.ID,'carVehicleTypes-menu-list')))
                
                # select the correct model variant
                selected_model_variant_element = select_model_variant()

                if selected_model_variant_element != None and selected_model_variant_element != "Unable to find car variant": # if a well matching car variant was found
                    selected_model_variant_element.click() # click the selected model variant
                else:
                    return "Unable to find car variant"


        # Open the webpage
        driver.get("https://my.tower.co.nz/quote/car/page1")    
        
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
                raise Exception("Unknown Error: Failed attempting to click business use button")

                
        
        # clicks the button to reset the input data (need to do this everytime except first time, as the page 'remembers' the previous iteration)
        try:
            Wait3.until(EC.presence_of_element_located( (By.ID, "vehicleUsedForBusiness-error-link"))).click()
        except exceptions.TimeoutException:
            pass


        # attempt to input the car registration number (if it both provided and valid)
        try: 
            enter_registration_number()

        except ValueError: # if the registration is invalid or not provided, then need to enter car details manually
            enter_car_details_manually()
        

        # enters whether the car has an immobiliser, if needed
        try:
            if data["Immobiliser"] == "Yes":
                Wait3.until(EC.element_to_be_clickable( (By.ID, "btnvehicleAlarm-0") )).click()
            else: # if No Immobiliser alarm
                Wait3.until(EC.element_to_be_clickable( (By.ID, "btnvehicleAlarm-1") )).click()
        except exceptions.TimeoutException:
            pass

        time.sleep(1) # wait for page to process information


        ## input the address the car is usually kept overnight at
        # inputing postcode + suburb
        Wait3.until(EC.element_to_be_clickable( (By.ID, "lnkManualAddress") )).click() # click this button to enter the address manually
        try:
            Wait3.until(EC.presence_of_element_located((By.ID, "txtAddress-street-number")) ).send_keys(data["Street_number"]) # input the street number
            Wait3.until(EC.presence_of_element_located((By.ID, "txtAddress-flat-number")) ).send_keys(data["Unit"]) # input the flat number (Unit number)
            Wait3.until(EC.presence_of_element_located((By.ID, "txtAddress-suburb-city-postcode")) ).send_keys(data["Street_name"]) # input the street name

            # find all addresses from the pop downs
            address_drop_down_options = Wait10.until(EC.presence_of_all_elements_located( (By.XPATH, f"//ul[@id='txtAddress-suburb-city-postcode-menu-list']/*/div/div[2]/div")))

            # creating the full address variable
            full_address = data["Street_name"] + ", " + data["Suburb"] + ", " + test_auto_data_df.loc[person_i, "City"] + " " + data["Postcode"]

            # intialising variables to allow us to find the address that matches the best with the spreadsheet data
            best_matching_address_option = None
            best_matching_address_score = 0

            for option in address_drop_down_options:
                # calculate a score that rates how good of a match the option is to the given input data
                score = fuzz.partial_ratio(full_address, option.text)

                # if a better match is found, save it
                if score > best_matching_address_score:
                    best_matching_address_score = score
                    best_matching_address_option = option

            # write the selected address to the output dataframe
            tower_output_df.loc[person_i, "TOWER_selected_address"] = f"Unit {data["Unit"]}, {data["Street_number"]} {best_matching_address_option.text}"

            # selecting the option
            best_matching_address_option.click()

        except exceptions.TimeoutException:
            raise Exception("Address Invalid")
            
        

        # enter the persons 'name' entering placeholder pseudonym either Jane or John Doe (If no name was provided)
        Wait10.until(EC.presence_of_element_located( (By.ID, "txtDriver-0-firstName") )).send_keys(data["First_name"])
        Wait3.until(EC.presence_of_element_located( (By.ID, "txtDriver-0-lastName") )).send_keys(data["Surname"])


        # enter the main drivers date of birth
        Wait3.until(EC.presence_of_element_located( (By.ID, "driverDob-0-day") )).send_keys(data["Birthdate_day"]) # input day
        Wait3.until(EC.presence_of_element_located( (By.ID, "driverDob-0-month") )).send_keys(data["Birthdate_month"]) # input month
        Wait3.until(EC.presence_of_element_located( (By.ID, "driverDob-0-year") )).send_keys(data["Birthdate_year"]) # input year

        try:
            time.sleep(0.5) # wait for page to load
            Wait3.until_not(EC.presence_of_element_located( (By.ID, "driver-dob-error") )) # checks if the person is old enough for tower to accept insuring them for the given car
        except: # if the warning is present, then this car/ person combo cannot be insured (the person is 'too yonug' for the given car)
            raise Exception("Website Does Not Quote For This Car Variant/ Person: 'We will only cover your vehicle when it is being driven by people aged 25 or over'")

        # select gender of main driver
        if data["Sex"] == "MALE":
            driver.find_element(By.ID, "btndriverGender-0-0").click() # clicking the 'Male' button
        else: #is Female
            driver.find_element(By.ID, "btndriverGender-0-1").click() # clicking the 'Female' button
        
        # input if there have been any indicents
        if data["Incidents_3_year"] > 0:
            driver.find_element(By.ID, "btndriverVehicleLoss-0-0").click() # clicks button saying that you have had an incident

            for i in range(data["Incidents_3_year"]):
                driver.find_element(By.ID, f"driverVehicleLossReason-0-{i}-toggle").click() # open incident type dropdown

                # Incident_type_index == ...; 1: "Broken windscreen", 2: "Collision", 6: "Theft"
                Wait3.until(EC.presence_of_element_located( (By.XPATH, f"//ul[@id='driverVehicleLossReason-0-{i}-menu-options']/li[{data["Incident_type_index"]}]") )).click() # find the correct incident then select it

                time.sleep(1) # wait for page to load


                # selects the correct year the indicent occured
                year_id_index = datetime.now().year - data[f"Incident{i+1}_year"] # gets the index for the id of the correct year element
                Wait3.until(EC.presence_of_element_located( (By.ID, f"btndriverVehicleLossReasonWhen-0-{i}-{year_id_index}") )).click()


                # if the incident type is a collision, damage while parked, or other causes
                if data["Incident_type_index"] == 2:

                    # if they had to pay an excess (if they made a claim)
                    if data["Incident_excess_paid"] == "Yes":
                        driver.find_element(By.ID, f"btndriverVehicleLossExcess-0-{i}-0").click() # click yes button
                    else:
                        driver.find_element(By.ID, f"btndriverVehicleLossExcess-0-{i}-1").click() # click no button

                # click the button to add an incident
                if i < data["Incidents_3_year"] - 1: # if not on the final incident, click button to 'Add another incident'
                    Wait3.until(EC.element_to_be_clickable((By.ID, f"mainIncidentsAccordion-0-{i}-add-link"))).click()

        else:
            driver.find_element(By.ID, "btndriverVehicleLoss-0-1").click() # clicks button saying that you have had no incidents


        # click button to specify how many additional drivers there are
        if data["Additional_drivers"]:
            # click button to add 1 more driver (WILL THROW AN ERROR AS WE CURRENTLY CANNOT POPULATE THE EXTRA DRIVER INFO)
            driver.find_element(By.ID, "mainDriversAccordion-0-add-link").click()


        # choose whether to exclude under all under 25 year old drivers from driving the car
        try:
            if data["Exclude_under_25"] == "Yes":
                driver.find_element(By.ID, "btnexcludeUnder25-0").click() # clicks yes button for exlcude under 25
            else: # select 'No'
                driver.find_element(By.ID, "btnexcludeUnder25-1").click() # clicks no button for exlcude under 25
        except:
            pass
        
        
        try:
            # try to click button 'Next: Customise' to move onto next page
            Wait3.until(EC.element_to_be_clickable( (By.ID, "btnSubmitPage") )).click()
        except exceptions.TimeoutException:
            # accept the privacy policy
            Wait3.until(EC.element_to_be_clickable( (By.ID, "privacyPolicy-label") )).click()
            
            # THEN click button 'Next: Customise' to move onto next page
            Wait3.until(EC.element_to_be_clickable( (By.ID, "btnSubmitPage") )).click()

            try:
                # check if "We're sorry The details you've given mean that we're currently unable to offer you cover." pop up has appeared
                Wait3.until(EC.presence_of_element_located((By.ID, "underwritingDialog")))

                # returning an error message
                raise Exception("Website Does Not Quote For This Car Variant/ Person")
            except exceptions.TimeoutException:
                pass

                

        # press button to aknowledge the extra excess
        try:
            Wait3.until(EC.element_to_be_clickable((By.ID, "btnClose"))).click()
            print("Extra Excess", end=" - ")
        except exceptions.TimeoutException:
            pass


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
        min_value = funct_defs.convert_money_str_to_int(Wait10.until(EC.presence_of_element_located( (By.XPATH, "//*[@id='agreedValueNew']/div[1]/div[1]/div[2]") )).text) # get the min agreed value
        max_value = funct_defs.convert_money_str_to_int(Wait10.until(EC.presence_of_element_located( (By.XPATH, "//*[@id='agreedValueNew']/div[1]/div[3]/div[2]") )).text) # get the max agreed value
        
        tower_output_df.loc[person_i, "TOWER_agreed_value_minimum"] = min_value # save the minimum allowed agreed value
        tower_output_df.loc[person_i, "TOWER_agreed_value_maximum"] = max_value # save the maximum allowed agreed value

        # check if our attempted agreed value is valid. if not, round up/down to the min/max value
        if data["Agreed_value"] > max_value:
            raise Exception("Invalid Input Data Error: AgreedValue Too High")
        elif data["Agreed_value"] < min_value:
            raise Exception("Invalid Input Data Error: AgreedValue Too Low")

        # output the corrected agreed value
        tower_output_df.loc[person_i, "TOWER_agreed_value"] = data["Agreed_value"]

        # inputs the agreed value input the input field (after making sure its valid)
        agreed_value_input = driver.find_element(By.ID, "agreedValueNewSliderField") # find the input field for the agreed value
        agreed_value_input.send_keys(Keys.CONTROL, "a") # select all current value
        agreed_value_input.send_keys(data["Agreed_value"]) # input the desired value, writing over the (selected) current value
        Wait3.until(EC.element_to_be_clickable((By.ID, "agreedValueNewSliderBtn"))).click() # click the 'Update agreed value' button


        time.sleep(1) # wait for page to load


        # input the persons desired level of excess
        Wait10.until(EC.element_to_be_clickable( (By.ID, f"btnexcess-{data["Excess_index"]}"))).click()

        # click button to state whether or not the car has any 'major' modifications
        if data["Modifications"] == "No":
            no_button = Wait10.until(EC.element_to_be_clickable( (By.XPATH, "//*[@id='btnaccessoriesOrModificationsDeclined-1']/div[2]/div") )) # find "No" modifications button
            driver.execute_script("arguments[0].scrollIntoView();", no_button) # scroll down until "No" modifications button is on screen (is needed to prevent the click from being intercepted)
            no_button.click() # click "No" modifications button
        else:
            raise Exception("Website Does Not Quote For This Car Variant/ Person: 'Not insurable: Modifications'")

        # click button to state whether or not the car has any 'minor' modifications
        if data["Modifications"] == "No":
            Wait10.until(EC.element_to_be_clickable( (By.ID, "btnaccessoriesOrModifications-1") )).click() # click "No" modifications button
        else:
            raise Exception("Website Does Not Quote For This Car Variant/ Person: 'Not insurable: Modifications'")
            

        # click 'Next: Summary' button to move onto the summary page
        next_button = Wait10.until(EC.element_to_be_clickable( (By.ID, "btnSubmitPage")))
        driver.execute_script("arguments[0].scrollIntoView();", next_button) # scroll down until "Next: Summary" button is on screen (is needed to prevent the click from being intercepted)
        Wait10.until(EC.element_to_be_clickable((By.ID, "btnSubmitPage"))).click()

        # attempt to click 'accept' button on the 'Sorry! Something's gone wrong' pop up
        try:
            Wait3.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="invalidVersionModal"]/footer/button'))).click()
        except:
            pass

        # move onto the next page "People"
        Wait10.until(EC.element_to_be_clickable((By.ID, "btnSubmitPage"))).click()


        # if a pop telling us 'an additional excess of ... will apply to claims of theft for this vehicle' appears, close it
        try:
            Wait10.until(EC.element_to_be_clickable((By.ID, "btnClose"))).click()
        except exceptions.TimeoutException: # if the popup not present then continue
            pass
        
        # click button to say that the driver has NOT "had their licence suspended or cancelled or had a special condition imposed"
        if data["License_suspended"].upper() == "NO":
            Wait10.until(EC.element_to_be_clickable((By.ID, "btndriver-0-licence-cancelled-1"))).click()
        else: # if they have had their license suspended or cancelled or had a special condition imposed
            raise Exception("Website Does Not Quote For This Car Variant/ Person: 'Cannot Insure Person Because License Suspended in Last 7 Years'")

        # click button to say that the policy will NOT be held by a business or trust
        if data["Business_or_trust"].upper() == "NO":
            Wait3.until(EC.element_to_be_clickable((By.ID, "btnownedByBusinessOrTrust-1"))).click()
        else:
            raise Exception("Website Does Not Quote For This Car Variant/ Person: 'Policy Owned by a Business or Trust'")


        # enter email address
        Wait3.until(EC.presence_of_element_located((By.ID, "txtDriver-0-email"))).send_keys(data["Email"])

        # enter phone number
        Wait3.until(EC.presence_of_element_located((By.ID, "txtDriver-0-phoneNumbers-0"))).send_keys(data["Phone_number"])

        # click button to go to next page 'Legal'
        Wait3.until(EC.element_to_be_clickable((By.ID, "btnSubmitPage"))).click()
        
        # click button to say I understand the 1st legal information declaration (the 'Yes' button)
        Wait10.until(EC.element_to_be_clickable((By.ID, "btnlegalDeclaration-0"))).click()

        # click button to say I understand the 2nd legal information (important things to call out) declaration (the 'Yes' button)
        Wait3.until(EC.element_to_be_clickable((By.ID, "btnexclusions-0"))).click()

        # click button to say I haven't had insurance refused or cancelled within the last 7 years (the 'No' button)
        if not data["Insurance_refused_7_years"]: # if person has never had insurance refused
            Wait3.until(EC.element_to_be_clickable((By.ID, "btninsuranceHistory-1"))).click()
        else:
            return "Website Does Not Quote For This Car Variant/ Person: Insurance Has Previously Been Refused"

        # click button to say I haven't had a claim declined or policy avoided in the last 7 years (the 'No' button)
        if not data["Claim_refused_7_years"]:
            Wait3.until(EC.element_to_be_clickable((By.ID, "btnclaimsDeclined-1"))).click()
        else:
            return "Website Does Not Quote For This Car Variant/ Person: Had claim refused within last 7 years"

        # click button to say I have not been convicted of Fraud, Arson, Bugulary, Wilfull damage, sexual offences, or drugs conviction within the last 7 years (the 'No' button)
        if not data["Crime_7_years"]:
            Wait3.until(EC.element_to_be_clickable((By.ID, "btncriminalHistory-1"))).click()
        else:
            return "Website Does Not Quote For This Car Variant/ Person: Commited Serious Crime within last 7 years"

        # select whether the car was purchased on finance
        if data["Finance_purchase"].upper() == "NO":
            Wait10.until(EC.element_to_be_clickable( (By.ID, "btnmoneyOwed-1") )).click() # click "No" Finance" button
        else:
            Wait10.until(EC.element_to_be_clickable( (By.ID, "btnmoneyOwed-0") )).click() # click "Yes" Finance" button
            Wait3.until(EC.presence_of_element_located((By.ID, "financialInterestedParty-0-financialInterestedParty-financial-institution-search"))).send_keys(data["Bank"]) # enter the finance provider as kiwibank
            Wait3.until(EC.element_to_be_clickable((By.XPATH, '//ul[@id="financialInterestedParty-0-financialInterestedParty-financial-institution-search-menu-list"]/li[1]'))).click() # select the first option as finance provider

        # click button to go to the next page 'Finalise'
        Wait3.until(EC.element_to_be_clickable((By.ID, "btnSubmitPage"))).click()

        # click button to say the person would not like to link an airpoints account (Click 'No' button)
        try:
            Wait3.until(EC.element_to_be_clickable((By.ID, "btnairpointsIsMember-1"))).click()
        except exceptions.TimeoutException:
            pass

        # input the desired start date
        start_date_input_element = Wait3.until(EC.presence_of_element_located((By.ID, "policyStartDatePicker")))
        start_date_input_element.send_keys(Keys.CONTROL + "a")
        start_date_input_element.send_keys(data["Policy_start_date"])

        # click button to move to next page ('Payment')
        Wait10.until(EC.element_to_be_clickable((By.ID, "btnSubmitPage"))).click()

        # scrape the monthly and yearly premiums
        monthly_premium = funct_defs.convert_money_str_to_int(Wait10.until(EC.presence_of_element_located( (By.ID, "btnpaymentFrequency-1"))).text, cents=True) # scrape the monthy premium and convert into an integer
        yearly_premium = funct_defs.convert_money_str_to_int(Wait10.until(EC.presence_of_element_located( (By.ID, "btnpaymentFrequency-2"))).text, cents=True) # scrape the yearly premium and convert into an integer


        # return the scraped premiums
        return round(monthly_premium, 2), round(yearly_premium, 2)


    # get time of start of each iteration
    start_time = time.time() 

    # run on the ith car/person
    try:
        # scrapes the insurance premium for a single vehicle and person
        tower_auto_premium = tower_auto_scrape_premium(tower_auto_data_format(person_i)) 

        # print the scraping results
        print(tower_auto_premium[0], tower_auto_premium[1], end =" -- ")

        # save the scraped premiums to the output dataset
        tower_output_df.loc[person_i, "TOWER_monthly_premium"] = tower_auto_premium[0] # monthly
        tower_output_df.loc[person_i, "TOWER_yearly_premium"] = tower_auto_premium[1] # yearly



    except Exception as error_message:
        
        # convert the error_message into a string
        error_message = str(error_message)

        # defining a list of known error messages
        errors_list = ["Website Does Not Quote For This Car Variant/ Person", "Unable to find car variant", "Invalid Input Data Error", "Several Car Variant Options Warning", "Excess cannot be changed from"]
        execute_bottom_code = True

        # checking if the error message is one of the known ones
        for error in errors_list:
            # checking if the error message that was returned is a known one
            if  error in error_message:
                print(error_message, end= " -- ")
                tower_output_df.loc[person_i, "TOWER_Error_code"] = error_message
                execute_bottom_code = False

        # if the error is not any of the known ones
        if execute_bottom_code:

            try:
                Wait3.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="invalidVersionModal"]/footer/button'))).click()
            except exceptions.TimeoutException:
                pass
            finally:
                print("Unknown Error!!", end = "--")
                tower_output_df.loc[person_i, "TOWER_Error_code"] = error_message

                # checking if the "More info required" pop-up has appeared (if it has then we cannot continue)
                try:
                    Wait3.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="recaptchaDialog"]/section/div/h5[contains(text(), "More info required!")]')))
                except:
                    pass
                else:
                    end_time = time.time() # get time of end of each iteration
                    print("Elapsed time:", round(end_time - start_time, 2)) # print out the length of time taken
                    return True # if "More info required" didn't appear, return True
                

    end_time = time.time() # get time of end of each iteration
    print("Elapsed time:", round(end_time - start_time, 2)) # print out the length of time taken

    # if "More info required" didn't appear, return False
    return False
    



def auto_scape_all():
    
    # define a variable that saves the number of rows to scrape
    num_rows_to_scrape = len(test_auto_data_df)
    
    # initialise this variable (counts the number of times 'More info required' pop-up appears in a row (if more than once the code stops and returns the current data))
    more_info_required_pop_up_counter = 0


    # loop through all cars in test spreadsheet
    for person_i in range(0, num_rows_to_scrape): 

        print(f"{person_i}: Tower: ", end = "") # print out the iteration number

        # set for this person, the PolicyStartDate to todays date
        test_auto_data_df.loc[person_i, "PolicyStartDate"] = date.today().strftime(format="%d/%m/%Y")
        tower_output_df.loc[person_i, "PolicyStartDate"] = datetime.strftime(date.today(), "%d/%m/%Y")

        # run on the ith car/person
        if tower_auto_scrape(person_i):
            more_info_required_pop_up_counter += 1

            # if the More info required pop-up appears more than once in a row, then return the so far scraped data
            if more_info_required_pop_up_counter > 1:
                funct_defs.export_auto_dataset(tower_output_df, "TOWER")
                print(f"'More info required!' pop-up appeared on the {person_i - 1} and {person_i} examples. Restart from first failed example using a different VPN location")
                return # exit this function

        else:
            # if this pop-up isnt the cause of two issues in a row
            more_info_required_pop_up_counter = 0

        # delete all cookies to reset the page
        try:
            driver.delete_all_cookies()
        except exceptions.TimeoutException: # if we timeout while trying to reset the cookies

                print("\n\nNew Webdriver window\n")
                driver.quit() # quit this current driver
                load_webdriver() # open a new webdriver session
    

    funct_defs.export_auto_dataset(tower_output_df, "TOWER")




def main():
    # performing all data reading in and preprocessing
    global test_auto_data_df, tower_output_df
    test_auto_data_df, tower_output_df = funct_defs.dataset_preprocess("TOWER")

    # loads chromedriver
    load_webdriver()

    # scrape all of the insurance premiums for the given cars from tower
    auto_scape_all()

    # Close the browser window
    driver.quit()

main()
# File path definitions
test_auto_data_xlsx = "C:\\Users\\samuel.johnston\\Documents\\Insurance_web-scraping\\test_auto_data1.xlsx"
#test_home_data_xlsx

# webscraping related imports
import time
from selenium.common import exceptions
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# data management/manipulation related imports
import pandas as pd
from datetime import datetime
import math

# defining a function that will scrape all of the aa cars
def aa_auto_scrape_all():
    # defining a function which take the information from the spreadsheet and formats it so it can be used to scrape premium from aa website
    def aa_auto_data_format(person_i):
        # getting the street address
        house_number = test_auto_data.loc[person_i,'Street_number']
        street_name = test_auto_data.loc[person_i,'Street_name']
        street_type = test_auto_data.loc[person_i,'Street_type']
        if "(" in test_auto_data.loc[person_i,'Street_name']:
            street_name = test_auto_data.loc[person_i,'Street_name'].split("(")[0].strip()

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

        # define a dict to store information for a given person and car for ami
        aa_data  = {"Cover_type":test_auto_data.loc[person_i,'CoverType'],
                    "AA_member":test_auto_data.loc[person_i,'AAMember'],
                    "Registration_number":test_auto_data.loc[person_i,'Registration'],
                    "Vehicle_year":test_auto_data.loc[person_i,'Vehicle_year'],
                    "Manufacturer":test_auto_data.loc[person_i,'Manufacturer'],
                    "Model":test_auto_data.loc[person_i,'Model'],
                    "Model_type":test_auto_data.loc[person_i,'Type'],
                    "Body_type":test_auto_data.loc[person_i,'Body'].upper(),
                    "Engine_size":"{}cc".format(int(test_auto_data.loc[person_i,'CC'])),
                    "Modifications":test_auto_data.loc[person_i,'Modifications'],
                    "Finance_purchase":test_auto_data.loc[person_i,'FinancePurchase'],
                    "Business_use":test_auto_data.loc[person_i,'BusinessUser'],
                    "Street_address": f"{house_number} {street_name} {street_type}",
                    "Suburb":test_auto_data.loc[person_i,'Suburb'].strip(),
                    "Postcode":test_auto_data.loc[person_i,'Postcode'],
                    "Birthdate_day":int(birthdate.strftime("%d")),
                    "Birthdate_month":birthdate.strftime("%B"),
                    "Birthdate_year":int(birthdate.strftime("%Y")),
                    "Sex":test_auto_data.loc[person_i,'Gender'],
                    "Incidents_3_year":Incidents_3_year,
                    "Current_insurer":current_insurer
                    }
        return aa_data
    

    # scrapes the insurance premium for a single vehicle/person at aa
    def aa_auto_scrape_premium(data):
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
            driver.find_element(By.XPATH, "//*[@id='aaMembershipDetails.aaMemberButtons']/label[1]/span").click()
        else:
            driver.find_element(By.XPATH, "//*[@id='aaMembershipDetails.aaMemberButtons']/label[2]/span").click()


        # attempt to input the car registration number (if it both provided and valid)
        try: 
            driver.find_element(By.ID, "vehicleRegistrationNumberNz").send_keys(data["Registration_number"]) # input registration number
            driver.find_element(By.ID, "vehicleRegistrationSearchButtonNz").click() # click check button

            # attempt to find the 1st option for car pop down (if present then we can continue)
            Wait.until(EC.visibility_of_element_located( (By.ID,  "vehicleDetailSummaryBoxAlt")))

        except: # if the registration is invalid or not provided, then need to enter car details manually
            pass
        
        # click button to move to car features page
        time.sleep(2)
        Wait.until(EC.element_to_be_clickable( (By.ID, "_eventId_submit") )).click()

        # click button to state whether or not the car has any modifications
        if data["Modifications"] == "No":
            Wait10.until(EC.element_to_be_clickable( (By.XPATH, "//*[@id='accessoriesAndModificationsButtons']/label[2]/span") )).click() # click "No" modifications button
        else:
            Wait10.until(EC.element_to_be_clickable( (By.XPATH, "//*[@id='accessoriesAndModificationsButtons']/label[1]/span") )).click() # click "Yes" modifications button
            raise Exception("Not insurable: Modifications")

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
            driver.find_element(By.ID, "address.suburbPostcodeRegionCity").send_keys(data['Postcode']) # input the postcode
            WebDriverWait(driver, 2).until(EC.element_to_be_clickable( (By.XPATH, "//*[@id='quote']/fieldset[3]/div[1]/div[1]/ul/li[contains(text(), '{}')]".format(data["Suburb"])) )).click() # click the pop down option with the correct subrub
        except exceptions.TimeoutException:
            driver.find_element(By.XPATH, "//*[@id='quote']/fieldset[3]/div[1]/div[1]/ul/li[1]").click() # if the above doesn't work, just select the first pop down option
        finally: # input street address
            driver.find_element(By.ID, "address.streetAddress").send_keys(data["Street_address"]) # inputting street address

        # click button to move to driver details page
        driver.find_element(By.ID, "_eventId_submit").click()

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
        Wait.until(EC.element_to_be_clickable( (By.XPATH, "//*[@id='previousInsurerList']//*[contains(text(), '{}')]".format(data["Current_insurer"])) )).click() # click the correct 'previous insurer'


        time.sleep(10)
    
    # loop through all cars in test spreadsheet
    for person_i in range(0, 1):#len(test_auto_data)):
        start_time = time.time() # get time of start of each iteration

        print(person_i, ": ", end = "")
        # run on the ith car/person
        #try:
        ami_auto_premium = aa_auto_scrape_premium(aa_auto_data_format(person_i))  # scrapes the insurance premium for a single vehicle and person at aa
        if ami_auto_premium != None: # if an actual result is returned
            monthy_premium, yearly_premium = ami_auto_premium[0], ami_auto_premium[1]
            print(monthy_premium, yearly_premium, end =" -- ")
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




# defining a function that will scrape all of the ami cars
def ami_auto_scrape_all():
    # defining a function which take the information from the spreadsheet and formats it so it can be used to scrape premium from ami website
    def ami_auto_data_format(person_i):
        # formatting street name and type into the correct format
        street_name = test_auto_data.loc[person_i,'Street_name']
        street_type = test_auto_data.loc[person_i,'Street_type']
        if "(" in test_auto_data.loc[person_i,'Street_name']:
            street_name = test_auto_data.loc[person_i,'Street_name'].split("(")[0].strip()

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
                    "Suburb":test_auto_data.loc[person_i,'Suburb'].strip(),
                    "Postcode":test_auto_data.loc[person_i,'Postcode'],
                    "Birthdate_day":int(birthdate.strftime("%d")),
                    "Birthdate_month":birthdate.strftime("%B"),
                    "Birthdate_year":int(birthdate.strftime("%Y")),
                    "Sex":test_auto_data.loc[person_i,'Gender'],
                    "Drivers_license_type":drivers_license_type,
                    "Drivers_license_years":drivers_license_years, # years since driver got their learners licence
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
    def ami_auto_scrape_premium(data):
        # Open the webpage
        driver.get("https://secure.ami.co.nz/css/car/step1")

        # attempt to input the car registration number (if it both provided and valid)
        registration_na = pd.isna(data["Registration_number"])
        if not registration_na: # if there is a registration number provided
            driver.find_element(By.ID, "vehicle_searchRegNo").send_keys(data["Registration_number"]) # input registration
            driver.find_element(By.ID, "ie_regSubmitButton").click() # click submit button

            # attempt to find the 1st option for car pop down (if present then we can continue)
            try: 
                Wait.until(EC.element_to_be_clickable( (By.ID,  "searchedVehicleSpan_0")))
                registration_invalid = False
            except: # if that element is not findable then the registration must have been invalid
                registration_invalid = True
        # is effectively an "else" statement for the above if
        if registration_na or registration_invalid: # if registration invalid or not provided we need to enter car details
            # Find the button "Make, Model, Year" and click it
            make_model_year_button = driver.find_element(By.ID, "ie_MMYPrepareButton")
            make_model_year_button.click()


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
                driver.find_element(By.ID, "Year").click() # find car year input box then open it
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
                print(data["Engine_size"])
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
            time.sleep(2) # wait for options list to pop up
            driver.find_element(By.XPATH, "//li[@class='ui-menu-item']//a[contains(text(), '{}') or contains(text(),'{}')]".format(data["Suburb"], data["Postcode"])).click()
        except: # if no pop up after inputting the street address, try inputting the suburb
            suburb_entry_element = driver.find_element(By.ID, "garagingAddress_manualSuburb")
            suburb_entry_element.send_keys(data["Suburb"])
            time.sleep(1) # wait for elements on the page to load
            try:
                driver.find_element(By.XPATH, "//li[@class='ui-menu-item']//a[contains(text(), '{}')]".format(data["Postcode"])).click() # try to find and click any pop down element that contains the postcode
            except:
                try: # try entering just the postcode into the suburb
                    suburb_entry_element.clear() # clears the textbox
                    suburb_entry_element.send_keys(data["Postcode"]) # type into the box just the postcode
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

        # delete all cookies to reset the page
        driver.delete_all_cookies()

        # wait until next page is loaded
        annual_risk_premium = Wait.until(EC.presence_of_element_located((By.ID, "annualRiskPremium")))

        # scrape the premium
        monthy_premium = float(driver.find_element(By.ID, "dollars").text.replace(",", "") + driver.find_element(By.ID, "cents").text)
        yearly_premium = float(annual_risk_premium.text.replace(",", "")[1:])

        # return the scraped premiums
        return monthy_premium, yearly_premium


    # loop through all cars in test spreadsheet
    for person_i in range(443, len(test_auto_data)):
        start_time = time.time() # get time of start of each iteration

        print(person_i, ": ", end = "")
        # run on the ith car/person
        try:
            ami_auto_premium = ami_auto_scrape_premium(ami_auto_data_format(person_i))
            if ami_auto_premium != None: # if an actual result is returned
                monthy_premium, yearly_premium = ami_auto_premium[0], ami_auto_premium[1]
                print(monthy_premium, yearly_premium, end =" -- ")
        except:
            try: # checks if the reason our code failed is because the 'we need more information' pop up appeareds
                Wait.until(EC.visibility_of_element_located( (By.XPATH, "//*[@id='ui-id-3' and text() = 'We need more information']") ) )
                print("Need more information", end= " -- ")
            except exceptions.TimeoutException:
                print("Unknown Error!!", end= " -- ")

        end_time = time.time() # get time of end of each iteration
        print("Elapsed time:", round(end_time - start_time,2)) # print out the length of time taken

def main():
    # reads in the test data for car insurance inputs
    global test_auto_data
    test_auto_data = pd.read_excel(test_auto_data_xlsx, dtype={"Postcode":"int"})
    
    # defines a function to reformat the postcodes in test_auto_data
    def postcode_reformat(postcode):
        postcode = str(postcode) # converts the postcode input into a string
        while len(postcode) != 4:
            postcode = f"0{postcode}"
        return postcode

    # pads out the front of postcodes with zeroes (as excel removes leading zeros)
    test_auto_data['Postcode'] = test_auto_data['Postcode'].apply( postcode_reformat ) 

    # setup the chromedriver options
    options = webdriver.ChromeOptions()
    options.add_experimental_option('excludeSwitches', ['enable-logging'])

    # loads chromedriver
    global driver # defines driver as a global variable
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))

    # define the implicit wait time for the session
    driver.implicitly_wait(1)

    # defines Wait and Wait10, dynamic wait templates
    global Wait
    Wait = WebDriverWait(driver, 5)

    global Wait10
    Wait10 = WebDriverWait(driver, 5)

    # scrape all of the insurance premiums for the given cars on aa
    aa_auto_scrape_all()

    # scrape all of the insurance premiums for the given cars on ami
    #ami_auto_scrape_all()

    # Close the browser window
    driver.quit()

main()
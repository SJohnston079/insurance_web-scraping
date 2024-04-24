# webscraping related imports
from selenium import webdriver
from selenium.common import exceptions
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait, Select
import time

# data management/manipulation related imports
import pandas as pd
from datetime import datetime, date
import re

# importing for more natural string comparisons
from fuzzywuzzy import fuzz


# importing several general functions (which are defined in the seperate python file called funct_defs)
import funct_defs

# defing the file path for test_home_data.csv (to allow us to read it in)
test_home_data_file_path = funct_defs.define_file_path()

"""
-------------------------
Useful functions
"""

# define a function to load the webdriver
def load_webdriver():
    # loads chromedriver
    global driver, Wait1, Wait3, Wait10
    driver, Wait1, Wait3, Wait10 = funct_defs.load_webdriver()

"""
-------------------------
"""

### Defining the key functions for scraping from Tower
################################################################################################################################

# defining a function to scrape from the given company website (for an individual person/house)
def ami_home_premium_scrape(person_i):
    
    # defining a function which take the information from the spreadsheet and formats it so it can be used to scrape premium from aa website
    def ami_home_data_format(person_i):
        # formatting the occupancy variable
        occupancy = test_home_data_df.loc[person_i,'Occupancy'].upper()


        # formatting current_renovations
        current_renovations = test_home_data_df.loc[person_i, 'AACurrentRenovations'].upper()

        if current_renovations == 'NON-STRUCTURAL>75K' or current_renovations =='STRUCTURAL': # if the renovations are "major"
            print("Webiste Does Not Quote For This House/ Person, Issue With AACurrentRenovations Column", end=" -- ")
            return "Webiste Does Not Quote For This House/ Person, Issue With AACurrentRenovations Column"
        else:
            # the current renovations are non-structural and are valued at least than $75,000, is True, else will be false (for answer 'No' in spreadsheet)
            current_renovations = current_renovations == 'NON-STRUCTURAL<75K'

        # formatting the address of the house, concatenate all the address elements eg num& Street name& type&... city..
        street_address = f"{test_home_data_df.loc[person_i,'Street_number']} {test_home_data_df.loc[person_i,'Street_name']} {test_home_data_df.loc[person_i,'Street_type']}"
        full_address = f"{street_address}, {test_home_data_df.loc[person_i, "Suburb"]}, {test_home_data_df.loc[person_i, "City"]}, {test_home_data_df.loc[person_i, "Postcode"]}"

        
        # formatting the building type
        building_type = test_home_data_df.loc[person_i, "BuildingType"].upper()
        if building_type in ["APARTMENT", "BOARDING HOUSE", "RETIREMENT UNIT"]:
            print("Webiste Does Not Quote For This House/ Person, Issue With BuildingType Column", end=" -- ")
            return "Webiste Does Not Quote For This House/ Person, Issue With BuildingType Column"
        elif building_type == "FREESTANDING HOUSE":
            building_type = "FreestandingHouse"
        elif "MULTI UNIT" in building_type: 
            building_type = "FlatOrUnit"
        elif building_type == "SEMI DETACHED / TOWNHOUSE":
            building_type = "SemiDetachedHouseOrTerrace"
        
        # formatting the contruction type
        construction_type = test_home_data_df.loc[person_i, "ConstructionType"].upper()
        if construction_type in ["ALUMINIUM", "BUTYNOL or MALTHOID", "CLADDING", "HARDIPLANK/HARDIFLEX"]:
            print("Webiste Does Not Quote For This House/ Person, Issue With ConstructionType Column", end=" -- ")
            return "Webiste Does Not Quote For This House/ Person, Issue With ConstructionType Column"
        elif construction_type == "STONE":
            construction_type = "NaturalStone"
        elif construction_type == "ROCKCOTE EPS":
            construction_type == "RockcoteEPS"
        elif construction_type == "CONCRETE":
            construction_type == "SolidConcreteWalls"
        else:
            construction_type = construction_type.title().replace(" ", "").replace("/", "")


        # formatting the roof type
        roof_type = test_home_data_df.loc[person_i, "RoofType"].upper()
        if roof_type in ["ALUMINIUM", "GLASS", "IRON (CORRUGATED)", "PLASTIC", "STEEL/COLOURBOND"]:
            print("Webiste Does Not Quote For This House/ Person, Issue With RoofType Column", end=" -- ")
            return "Webiste Does Not Quote For This House/ Person, Issue With RoofType Column"
        else:
            roof_type = roof_type.title().replace(" ", "").replace("/", "")


        # formatting the number of stories in the house
        num_stories = test_home_data_df.loc[person_i, "NumberOfStories"]
        if num_stories > 2:
            return "Webiste Does Not Quote For This House/ Person, Issue With NumberOfStories Column"
        elif num_stories > 0: # if either 1 or 2 stories in the house
            num_stories = str(num_stories)
        else: # if num stories < 1, which is invalid as it is not possible for a house to have less than 1 story and still be a house
            return "Invalid Input Data Error: NumberOfStories Column"

        # define a dict to store information for a given house/person
        data  = {"Occupancy":occupancy, 
                "Current_renovations":current_renovations,
                 "Street_address":street_address,
                 "Full_address":full_address,
                 "Birthdate":str(test_home_data_df.loc[person_i,'DOB'].strftime('%d/%m/%Y')),
                 "Building_type":building_type,
                 "Construction_type":construction_type,
                 "Roof_type":roof_type,
                 "Num_stories":num_stories
                }  
        return data


    # scrapes the insurance premium for a single vehicle/person at aa
    def ami_home_scrape_premium(data):

        # Opens the AMI House Insurance webpage
        driver.get("https://my.tower.co.nz/quote/bundle-builder")

           ## selecting the correct type of house insurance
        if data["Occupancy"] in ['UNOCCUPIED', 'BOARDING HOUSE']: # if the house is unoccupied or is a boarding house
            # if the house is unoccupied or is a boarding house, we must call in to get cover (so cannot be scraped)
            return "Webiste Does Not Quote For This House/ Person, Issue With Occupancy Column" 
        
        elif data["Occupancy"] in ['OWNER OCCUPIED', 'HOLIDAY HOME']: # if can be covered by standard home insurance
            # selects 'Premier House' (clicks 'Get a quote button')
            Wait10.until(EC.element_to_be_clickable((By.ID, 'product-picker-item-house-add-button'))).click() 

        elif data["Occupancy"] in ['RENTED', 'LET TO FAMILY/EMPLOYEES']:
            # selects 'Premier Rental Property (Landlord Insurance)' (clicks 'Get a quote button')
            Wait10.until(EC.element_to_be_clickable((By.ID, 'product-picker-item-landlord-add-button'))).click() 
        else:
            print(f"Data Entry ERROR - Occupancy {test_home_data_df.loc[person_i,'Occupancy']} not understood", end=' -- ')
            return "Invalid Input Data Error: Occupancy Column"

        # if already a tower customer
        if test_home_data_df.loc[person_i, "HaveTowerPolicies"].upper() == "NO":
            Wait10.until(EC.element_to_be_clickable((By.ID, 'btnexistingCustomer-1'))).click()
        elif test_home_data_df.loc[person_i, "HaveTowerPolicies"].upper() == "YES":
            Wait10.until(EC.element_to_be_clickable((By.ID, 'btnexistingCustomer-0'))).click()
        else:
            print(f"Data Entry ERROR - HaveTowerPolicies {test_home_data_df.loc[person_i,'HaveTowerPolicies']} not understood", end=' -- ')
            return "Invalid Input Data Error: HaveTowerPolicies Column"
        
        #begin the quote
        driver.find_element(By.ID, 'btnSubmitPage').click()

         ## entering the house address
        # entering suburb + postcode + city
        Wait3.until(EC.element_to_be_clickable((By.ID, "txtAddress-address-search"))).send_keys(data["Street_address"]) # entering the street address into the input box
        time.sleep(3) # wait for the page to load
        address_options = driver.find_elements(By.XPATH, "//*[@id='ui-id-2']/li/a")

        # Finds the option with the highest score. A higher score means the option is more similar to the data from the database (as represented by the Full_address string)
        best_match_option = max(address_options, key=lambda option: fuzz.ratio(data["Full_address"], option.text))
        output_df.loc[person_i, "AMI_selected_address"] = best_match_option.text # outputting the selected address
        best_match_option.click() # selecting the highest score option

        time.sleep(100000) # wait for testing purposes

        ## scraping the home premiums
        monthly_premium = Wait10.until(EC.presence_of_element_located((By.XPATH, '//*[@id="mqs-small"]/div/div[1]/div[1]/span'))).text
        yearly_premium = driver.find_element(By.XPATH, '//*[@id="mqs-small"]/div/div[1]/div[3]/span[1]').text

        # reformatting the montly and yearly premiums into integers
        monthly_premium, yearly_premium = funct_defs.convert_money_str_to_int(monthly_premium, cents=True), funct_defs. convert_money_str_to_int(yearly_premium, cents=True)

        # returning the monthly/yearly premium and the adjusted agreed value
        return monthly_premium, yearly_premium


    # get time of start of each iteration
    start_time = time.time() 


    # run on the ith car/person
    try:
        # scrapes the insurance premium for a single house/person
        home_premiums = ami_home_scrape_premium(ami_home_data_format(person_i))

        if home_premiums != None and not isinstance(home_premiums, str): # if home_premiums is the scraped premiums (NOT just an error message)

            # print the scraping results
            print(home_premiums[0], home_premiums[1], end =" -- ")

            # save the scraped premiums to the output dataset
            output_df.loc[person_i, "AMI_monthly_premium"] = home_premiums[0] # monthly
            output_df.loc[person_i, "AMI_yearly_premium"] = home_premiums[1] # yearly

        # all these are processing error codes
        elif "Webiste Does Not Quote For This House/ Person" in home_premiums:
            output_df.loc[person_i, "AMI_Error_code"] = home_premiums
        elif "Invalid Input Data Error" in home_premiums:
            output_df.loc[person_i, "AMI_Error_code"] = home_premiums
        else:
            output_df.loc[person_i, "AMI_Error_code"] = "Unknown Error"

    except:
        try: # checks if the reason our code failed is because the 'we need more information' pop up appeareds
            Wait3.until(EC.visibility_of_element_located( (By.XPATH, "//*[@id='ui-id-3' and text() = 'We need more information']") ) )
            print("Need more information", end= " -- ")
            output_df.loc[person_i, "AMI_Error_code"] = "Webiste Does Not Quote For This House/ Person"
        except exceptions.TimeoutException:
            raise Exception("View Errors")
            print("Unknown Error!!", end= " -- ")
            output_df.loc[person_i, "AMI_Error_code"] = "Unknown Error"


    end_time = time.time() # get time of end of each iteration
    print("Elapsed time:", round(end_time - start_time, 2)) # print out the length of time taken
    

### a function to scrape premiums all given examples on aa's website
def aa_auto_scape_all():
    # define which row incidies to scrape from
    input_indexes = funct_defs.read_indicies_to_scrape()

    # loop through all examples in test_home_data spreadsheet
    for person_i in input_indexes: 

        print(f"{person_i}: AMI: ", end = "") # print out the iteration number

        # set for this person, the PolicyStartDate to todays date
        test_home_data_df.loc[person_i, "PolicyStartDate"] = datetime.strftime(date.today(), "%d/%m/%Y")

        # run on the ith car/person
        ami_home_premium_scrape(person_i)
        

        # delete all cookies to reset the page
        try:
            driver.delete_all_cookies()
        except exceptions.TimeoutException: # if we timeout while trying to reset the cookies

                print("\n\nNew Webdriver window\n")
                driver.quit() # quit this current driver
                load_webdriver() # open a new webdriver session


    #funct_defs.export_auto_dataset(input_indexes)




def main():
    # performing all data reading in and preprocessing
    global test_home_data_df, output_df
    test_home_data_df, output_df = funct_defs.dataset_preprocess("AMI")

    # loading the webdriver
    load_webdriver()

    # scrape all of the insurance premiums for the given cars from aa
    aa_auto_scape_all()

    # Close the browser window
    #driver.quit()

main()
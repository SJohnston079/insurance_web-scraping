# webscraping related imports
from selenium import webdriver
from selenium.common import exceptions
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.common.keys import Keys
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


## for processing bank codes
# listing all the bank codes and full names
bank_codes = {
    "ABL": "ASB Bank Limited",
    "ABNZ": "ANZ Bank New Zealand Limited",
    "BNZ": "Bank of New Zealand",
    "NZCU": "NZCU",
    "GEMO": "GE Money",
    "HNZL": "The Hongkong and Shanghai Banking Corporation Limited",
    "KWBL": "Kiwibank Limited",
    "SCBS": "Southern Cross Building Society",
    "PSIS": "Co-Operative Bank",
    "TSBL": "TSB Bank Limited",
    "AAF": "AA Finance",
    "AAFL": "AA Finance Limited (AA Money)",
    "MHTC": "Mortgage Holding Trust Company Limited",
    "NZLL": "New Zealand Home Lending Limited",
    "NZHL": "NZ Home Loans",
    "OTHR": "Other (Not Listed)",
    "SBS": "Southland Building Society (SBS Bank)",
    "SOVE": "Sovereign",
    "WNZL": "Westpac New Zealand Limited"
}

# function to return the full name when given a code
def get_bank_name(code: str) -> str:
    """
    Returns the long name of a bank given its code.

    :param code: The bank code
    :type code: str
    :return: The long name of the bank
    :rtype: str
    """
    return bank_codes.get(code.upper(), "Unknown bank")


"""
-------------------------
"""

### Defining the key functions for scraping from Tower
################################################################################################################################

# defining a function to scrape from the given company website (for an individual person/house)
def tower_home_premium_scrape(person_i, first_attempt):
    
    # defining a function which take the information from the spreadsheet and formats it so it can be used to scrape premium from tower website
    def tower_home_data_format(person_i):
        # formatting the occupancy variable
        occupancy = test_home_data_df.loc[person_i,'Occupancy'].upper()

        if occupancy == "HOLIDAY HOME" and test_home_data_df.loc[person_i,"OwnerStayInHouseWithinNextYear"].upper() == "NO":
            raise Exception("Webiste Does Not Quote For This House/ Person, Issue With OwnerStayInHouseWithinNextYear Column")



        # formatting the address of the house, concatenate all the address elements eg num& Street name& type&... city..
        street_address = f"{test_home_data_df.loc[person_i,'Street_number']} {test_home_data_df.loc[person_i,'Street_name']} {test_home_data_df.loc[person_i,'Street_type']}"
        full_address = f"{street_address}, {test_home_data_df.loc[person_i, "Suburb"]}, {test_home_data_df.loc[person_i, "City"]}, {test_home_data_df.loc[person_i, "Postcode"]}"


        # formatting the contruction type
        wall_material = test_home_data_df.loc[person_i, "ConstructionType"].upper()


        if wall_material in ["BRICK VENEER", "DOUBLE BRICK", "MUD BRICK", "STONEWORK SOLID","STUCCO"]:
            wall_material = wall_material[0] + wall_material[1:].lower()
        elif wall_material in ["CONCRETE SOLID", "VINYL CLADDING"]:
            wall_material = "Other"
        elif wall_material == "CONCRETE BLOCK":
            wall_material = "Blockwork"
        elif wall_material in ["FIBRE CEMENT CLADDING", "HARDIPLANK/HARDIFLEX"]:
            wall_material = "Artificial weatherboard/plank cladding"
        elif wall_material in ["METAL CLADDING", "ALUMINIUM"]:
            wall_material = "Sheet cladding"
        elif wall_material == "ROCKCOTE EPS":
            wall_material = "Rockcote EPS"
        elif wall_material == "NATURAL STONE CLADDING":
            wall_material = "Stonework veneer"
        elif wall_material == "TIMBER / WEATHERBOARD":
            wall_material = "Weatherboard plank cladding" 
        else:
            raise Exception("Invalid Input Data Error: ConstructionType Column")


        ## formatting the roof material type
        roof_material = test_home_data_df.loc[person_i, "RoofType"].upper()

        # roof type formatting options
        if roof_material in ["CONCRETE SOLID", "TIMBER"]:
            roof_material = "Other"

        elif  roof_material in ["CEMENT TILES", "CONCRETE TILES"]:
            roof_material = "Pitched concrete tiles"

        elif "FLAT-FIBRO" in roof_material:
            roof_material = "Flat fibre cement"

        elif "PITCHED-FIBRO" in roof_material:
            roof_material = "Pitched fibre cement covering"
        
        elif roof_material == "SLATE":
            roof_material = "Pitched slate"

        elif roof_material == "TERRACOTTA/CLAY TILES":
            roof_material = "Pitched terracotta tiles"

        elif roof_material == "SHINGLES":
            roof_material = "Pitched timber shingles"

        elif roof_material == "FLAT-MEMBRANE":
            roof_material = "Flat membrane"

        elif "ALUMINIUM" in roof_material or "IRON" in roof_material or "STEEL" in roof_material: # if the material is 'metal'

            # find out if the metal roof is flat or pitched
            roof_angle = roof_material.split("-")[0]
            if roof_angle == "FLAT":
                roof_material = "Flat metal covering"
            elif roof_angle == "PITCHED":
                roof_material = "Pitched metal covering"
            else:
                raise Exception("Invalid Input Data Error: RoofType Column")
            
        else:
            raise Exception("Invalid Input Data Error: RoofType Column")


        # formatting the number of stories in the house
        num_stories = test_home_data_df.loc[person_i, "NumberOfStories"]
        if num_stories > 2:
            raise  Exception("Webiste Does Not Quote For This House/ Person, Issue With NumberOfStories Column")
        elif num_stories > 0: # if either 1 or 2 stories in the house
            num_stories = str(num_stories)
        else: # if num stories < 1, which is invalid as it is not possible for a house to have less than 1 story and still be a house
            raise Exception("Invalid Input Data Error: NumberOfStories Column")
        

        # handling whether any business is run at the house
        business_use = test_home_data_df.loc[person_i, "BusinessUser"].upper()
        if business_use in ["Yes-Hobby farm","Yes-Other", "Yes->50% of house used for business"]:
            raise Exception("Webiste Does Not Quote For This House/ Person, Issue With BusinessUser Column")
        elif not( business_use == "NO" or "HOME OFFICE" in business_use or "BED AND BREAKFAST" in business_use): # if the value is not one of these, then we don't understand what the value in the BusinessUser column so raise an error
            raise Exception("Invalid Input Data Error: BusinessUser Column")


        # formatting whether the house/person suffered loss or damage to a house within the last 3 years?
        if funct_defs.check_date_range(test_home_data_df.loc[person_i, "Date_of_incident"], 3):
            type_incident = test_home_data_df.loc[person_i, "Type_incident"] # if incident happened within lat 3 years, save the incident type
            
            if type_incident == "Accidental Glass Breakage":
                type_incident = "Broken glass"

            elif type_incident in ["Accidental Loss/Damage at the home", "Accidental Loss/Damage away from home"]:
                type_incident = "Accidental Damage"

            elif type_incident == "Escaped water or other liquid":
                type_incident = "Water damage"

            elif type_incident == "Malicious damage/Vandalism":
                type_incident = "Malicious damage"

            elif type_incident == "Motor burnout, fusion or food spoilage":
                type_incident = "Fusion"

            elif type_incident == "Natural Event":
                type_incident = "Earthquake"

            elif type_incident in ["Damage by an animal", "Explosion", "Impact (e.g. falling tree, space debris or vehicle)", "Legal Liability", "Loss of Rent", "Other"]:
                type_incident = "Other"

            elif type_incident == "Storm, cyclone or rainwater runoff":
                type_incident = "Storm damage"

            elif type_incident in ["Burglary (with break in)", "Fire", "Flood at current address", "Flood at previous address", "Theft (without break in)"]:
                raise Exception("Webiste Does Not Quote For This House/ Person, Issue With Type_incident Column")
            else:
                raise Exception("Invalid Input Data Error: Type_incident Column")
        else: 
            type_incident = "No Incident" # if incident happened longer than 3 years ago, say that no incident occured 


        # handling whether the house is well maintained
        if test_home_data_df.loc[person_i, "HouseWellMaintained"].upper() == "YES":
            pass

        elif test_home_data_df.loc[person_i, "HouseWellMaintained"].upper() == "NO":
            raise Exception("Webiste Does Not Quote For This House/ Person, Issue With HouseWellMaintained Column")
        
        else:
            raise Exception("Invalid Input Data Error: Issue with HouseWellMaintained Column")


        # handling whether the person has had insurance refused within last 7 years
        if test_home_data_df.loc[person_i, "Insurance Refused In Last 7 Years"].upper() == "NO":
            pass

        elif "YES" in test_home_data_df.loc[person_i, "Insurance Refused In Last 7 Years"].upper():
            raise Exception("Webiste Does Not Quote For This House/ Person, Issue With 'Insurance Refused In Last 7 Years' Column")
        
        else:
            raise Exception("Invalid Input Data Error: Issue with 'Insurance Refused In Last 7 Years' Column")
        

        # handling whether the person has commited a serious crime (Fraud, Arson, Burglary or Theft, Wilful Damage, Sexual Offence, Drugs Conviction (other than cannabis possession) ) within the last 7 years
        if test_home_data_df.loc[person_i, "Crime in Last 7 Years"].upper() == "NO":
            pass

        elif test_home_data_df.loc[person_i, "Crime in Last 7 Years"].upper() == "YES":
            raise Exception("Webiste Does Not Quote For This House/ Person, Issue With 'Crime in Last 7 Years' Column")
        
        else:
            raise Exception("Invalid Input Data Error: Issue with 'Crime in Last 7 Years' Column")
        

        # preprocessing whether the person has any finance on the house they are insuring
        mortgage_bank = test_home_data_df.loc[person_i, "Mortgage"].upper()

        if "YES" in mortgage_bank:
            mortgage_bank = mortgage_bank.replace("YES-", "") # filtering out the "YES-" from the from so we just have the bank code
            mortgage_bank = get_bank_name(mortgage_bank) # getting the full name from the bank code

        elif mortgage_bank == "NO":
            mortgage_bank = "None"
        else:
            raise Exception("Invalid Input Data Error: Issue with Mortgage Column")

        # define a dict to store information for a given house/person
        data  = {"Occupancy":occupancy, 
                 "Street_address":street_address,
                 "Full_address":full_address,
                 "Birthdate_day":str(test_home_data_df.loc[person_i,'DOB'].strftime('%d')),
                 "Birthdate_month":str(test_home_data_df.loc[person_i,'DOB'].strftime('%m')),
                 "Birthdate_year":str(test_home_data_df.loc[person_i,'DOB'].strftime('%Y')),
                 "Wall_material":wall_material,
                 "Roof_material":roof_material,
                 "Type_incident":type_incident,
                 "Incident_year":test_home_data_df.loc[person_i, "Date_of_incident"].strftime("%Y"),
                 "Mortgage_bank":mortgage_bank
                }  
        return data


    # scrapes the insurance premium for a single vehicle/person at tower
    def tower_home_scrape_premium(data):
        def initialise_tower_website():
            ## selecting the correct type of house insurance
            if data["Occupancy"] =='UNOCCUPIED': # if the house is unoccupied or is a boarding house
                # if the house is unoccupied or is a boarding house, we must call in to get cover (so cannot be scraped)
                raise Exception("Webiste Does Not Quote For This House/ Person, Issue With Occupancy Column")
            
            elif data["Occupancy"] in ['OWNER OCCUPIED', 'HOLIDAY HOME']: # if can be covered by standard home insurance
                # selects 'Premier House' (clicks 'Get a quote button')
                Wait10.until(EC.element_to_be_clickable((By.ID, 'product-picker-item-house-add-button'))).click() 

            elif data["Occupancy"] in ['RENTED', 'LET TO FAMILY', 'LET TO EMPLOYEES']:
                # selects 'Premier Rental Property (Landlord Insurance)' (clicks 'Get a quote button')
                Wait10.until(EC.element_to_be_clickable((By.ID, 'product-picker-item-landlord-add-button'))).click() 
            else:
                raise Exception("Invalid Input Data Error: Occupancy Column")

            
            time.sleep(1) # wait for the page to load


            # clicking the button to say the person is not already a tower customer (so we don't have to log in)
            element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, 'btnexistingCustomer-1'))) # finding the element to click
            driver.execute_script("arguments[0].scrollIntoView();", element) # Scroll to the element
            element.click() # Click the element


            time.sleep(1) # wait for the page to load


            #begin the quote
            element = Wait10.until(EC.element_to_be_clickable((By.ID, 'btnSubmitPage')))
            driver.execute_script("arguments[0].scrollIntoView();", element)
            element.click() # Click the element


            time.sleep(5) # wait for the page to load


        # Opens the Tower House Insurance webpage
        driver.get("https://my.tower.co.nz/quote/bundle-builder")

        # open the website and navigate to the correct page (either HomeOwners or Landlords insurance)
        initialise_tower_website()


         ## entering the house address
        # entering suburb + postcode + city
        Wait10.until(EC.element_to_be_clickable((By.ID, "txtAddress-address-search"))).send_keys(data["Street_address"]) # entering the street address into the input box


        # wait for the page to load
        time.sleep(3) 


        # verifying that an error doesn't appear after entering in the address
        try:
            Wait3.until(EC.visibility_of_element_located((By.CLASS_NAME, 'error-container')))
        except exceptions.TimeoutException:
            pass
        else:
            raise Exception("Webiste Does Not Quote For This House/ Person: Issue With 'The Address of the House' (Want you to call in)")

        # finding all of the dropdown address options
        address_options = driver.find_elements(By.XPATH, '//*[@id="txtAddress-address-search-menu-list"]/li/div/div[2]/div')
        
        # Finds the option with the highest score. A higher score means the option is more similar to the data from the database (as represented by the Full_address string)
        best_match_option = max(address_options, key=lambda option: fuzz.ratio(data["Full_address"], option.text))
        tower_output_df.loc[person_i, "Tower_selected_address"] = best_match_option.text # outputting the selected address
        best_match_option.click() # selecting the highest score option
        
        # checking for issues with the address that was input
        try: 
            Wait3.until(EC.visibility_of_element_located((By.ID, 'address-search-error-hAction'))) # checks for error message that says "It looks like your property is a body corporate. Based on the details you've given, we're currently unable to offer you cover. If this is wrong, please call us on 0800 370 068."
        except:
            pass
        else:
            raise Exception("Webiste Does Not Quote For This House/ Person: Issue With 'The Address of the House' (They think it is a part of a body corporate (can call in to dispute))")


        ## enetering details about the house
        # entering the year the house was built
        year_built_input = Wait10.until(EC.visibility_of_element_located((By.ID, 'yearBuilt')))
        year_built_input.clear() # removes the current content of the text box
        year_built_input.send_keys(test_home_data_df.loc[person_i, "YearBuilt"])
        

        # inputting the number of levels the house has
        Wait1.until(EC.element_to_be_clickable((By.ID, 'numberOfStories-toggle'))).click()
        Wait1.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="numberOfStories-menu-options"]/li/div/div[2]/div'))).click()
        
        # entering the floor area of the house in square metres
        house_floor_area_input = Wait1.until(EC.visibility_of_element_located((By.ID, 'livingArea'))) 
        house_floor_area_input.clear() # removes the current content of the text box
        house_floor_area_input.send_keys(test_home_data_df.loc[person_i, "DwellingFloorArea"])

        # checking that the floor area of the house is not invalid
        try:
            Wait1.until(EC.visibility_of_element_located((By.ID, 'living-area-error')))
        except exceptions.TimeoutException:
            pass
        else:
            raise Exception("Webiste Does Not Quote For This House/ Person: DwellingFloorArea Column: (Too large, so they want us to call in)")
        

        # entering the floor area of all of the outbuildings in square metres (Outbuildings are structures that are not connected to the main house. These include garages, sheds and greenhouses, but not sleepouts or granny flats)
        outbuildings_floor_area_input = Wait1.until(EC.visibility_of_element_located((By.ID, 'floorArea')))
        outbuildings_floor_area_input.clear() # removes the current content of the text box
        outbuildings_floor_area_input.send_keys(test_home_data_df.loc[person_i, "GarageFloorArea"])

        # checking that the outbuilding floor area is not invalid
        try:
            Wait1.until(EC.visibility_of_element_located((By.ID, 'floor-area-error')))
        except exceptions.TimeoutException:
            pass
        else:
            raise Exception("Webiste Does Not Quote For This House/ Person: GarageFloorArea Column: (Total outbuilding area is too high, so they want us to call in)")
        

        # inputting the main wall material type
        Wait1.until(EC.element_to_be_clickable((By.ID, 'constructionTypeCd-toggle'))).click()
        time.sleep(1) # wait for page to loadf
        Wait3.until(EC.element_to_be_clickable((By.XPATH, f'//*[@id="constructionTypeCd-menu-options"]/li/div/div[2]/div[contains(text(), "{data["Wall_material"]}")]'))).click()

        # inputting the main roof material type
        Wait1.until(EC.element_to_be_clickable((By.ID, 'roofTypeCd-toggle'))).click()
        time.sleep(1) # wait for page to load
        Wait3.until(EC.element_to_be_clickable((By.XPATH, f'//*[@id="roofTypeCd-menu-options"]/li/div/div[2]/div[contains(text(), "{data["Roof_material"]}")]'))).click()

        # inputting the slope of the section that the house is on
        Wait1.until(EC.element_to_be_clickable((By.ID, 'slope-toggle'))).click()
        if test_home_data_df.loc[person_i, "LandShape"] == "FlatAndGentle": # flat or gentle slope
            Wait3.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="slope-menu-options"]/li[1]'))).click() # click first option "Flat or gentle slope (up to about 5 degrees)"

        elif test_home_data_df.loc[person_i, "LandShape"] == "Moderate": # moderately steep slope
            Wait3.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="slope-menu-options"]/li[2]'))).click() # click second option "Moderate slope (about 15 degrees)"

        elif test_home_data_df.loc[person_i, "LandShape"] == "Severe": # very steep slope
            Wait3.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="slope-menu-options"]/li[3]'))).click() # click third option "Severe slope (35 degrees or more)"

        else: # data input error
            raise Exception("Invalid Input Data Error: ConstructionType Column")
        

        # house construction quality
        Wait1.until(EC.element_to_be_clickable((By.ID, "quality-toggle"))).click()
        if test_home_data_df.loc[person_i, "HouseStandard"] == "Ordinary": # select the 'standard' construction quality option
            Wait3.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="quality-menu-options"]/li[1]'))).click() 

        elif test_home_data_df.loc[person_i, "HouseStandard"] == "Quality": # select the 'high' construction quality option
            Wait3.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="quality-menu-options"]/li[2]'))).click() 

        elif test_home_data_df.loc[person_i, "HouseStandard"] == "Prestige": # select the 'prestige' construction quality option
            Wait3.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="quality-menu-options"]/li[3]'))).click() 

        else: # data input error
            raise Exception("Invalid Input Data Error: ConstructionType Column")
        

        # click the button to confirm the details are correct
        driver.find_element(By.ID, "infoCorrect").click()

        # entering the sum insured value
        sum_insured_input_box = Wait10.until(EC.presence_of_element_located((By.ID, "sumInsured")))
        sum_insured_input_box.send_keys(Keys.CONTROL + "a") # select the whole current sum insured input box
        sum_insured_input_box.send_keys(Keys.DELETE) # delete all current content in sum insured input box (all content we just selected)
        sum_insured_input_box.send_keys(str(test_home_data_df.loc[person_i, "SumInsured"]))
        driver.find_element(By.ID, 'btnSumInsuredUpdate').click() # click button to confirm sum insured value
        time.sleep(1) # wait for the page to load


        # scraping the estimated sum insurance value from the website
        try:
            tower_output_df.loc[person_i, "Tower_estimated_replacement_cost"] = round(funct_defs.convert_money_str_to_int(driver.find_element(By.XPATH, '//*[@id="questionSumInsuredAmount"]/div[5]', cents=True).get_attribute("innerHTML")), 2)
        except: # if there is no estimate of the sum insured value
            tower_output_df.loc[person_i, "Tower_estimated_replacement_cost"] = "No Estimate Provided"
        else:
            # checking to see if the sum insured value is acceptable
            try:
                Wait3.until(EC.visibility_of_element_located((By.ID, 'sum-insured-underwriting'))) # checking if the error message occurs

            except exceptions.TimeoutException: # if can't find any error message pop-up
                pass

            else: # if we do find an error message pop up
                # if the sum insured value is larger than the Tower_estimated_cost, then it must be too large (as we can only get to this part if we are on the error page (we check it above))
                if test_home_data_df.loc[person_i, "SumInsured"] > tower_output_df.loc[person_i, "Tower_estimated_replacement_cost"]:
                    raise Exception("Webiste Does Not Quote For This House/ Person: SumInsured Column: Value Too Large")
                else: # if the sum insured value is smaller than the Tower_estimated_cost, then it must be too small (as we can only get to this part if we are on the error page (we check it above))
                    raise Exception("Webiste Does Not Quote For This House/ Person: SumInsured Column: Value Too Small")


        
        # check if the question "Has the house ever been re-roofed, re-lined and re-wired?" appears
        try:  
            # answer 'Yes' ('No' case is handled in data formatting)
            Wait3.until(EC.element_to_be_clickable((By.ID, 'btnreroofedRelinedRewired-0'))).click()

        except exceptions.TimeoutException:
            pass


        # answer whether the house or land ever been identified as at risk from a natural hazard?
        if test_home_data_df.loc[person_i, "HazardRisk"].upper() == "NO":
            Wait3.until(EC.element_to_be_clickable((By.ID, 'btnnaturalHazard-1'))).click() # click 'No' Button

        elif test_home_data_df.loc[person_i, "HazardRisk"].upper() == "YES":
            raise Exception("Webiste Does Not Quote For This House/ Person: Due to HazardRisk Column")
        
        else: # if the input is not recognisable
            raise Exception("Invalid Input Data Error: HazardRisk Column")


        # answer whether the contains any seperate self contained dwellings
        if test_home_data_df.loc[person_i, "NumSelfContainedDwellings"] == 1:
            Wait3.until(EC.element_to_be_clickable((By.ID, 'btnexternalSelfContainedUnit-1'))).click() # click 'No' Button

        elif test_home_data_df.loc[person_i, "NumSelfContainedDwellings"] > 1:
            raise Exception("Webiste Does Not Quote For This House/ Person: Due to NumSelfContainedDwellings Column")
        
        else: # if the input is not recognisable
            raise Exception("Invalid Input Data Error: NumSelfContainedDwellings Column")


        # select the correct occupancy situation at the house
        if data["Occupancy"] == "OWNER OCCUPIED":

            # checking if there are paying renters or boaders who live with the owner at the property
            if test_home_data_df.loc[person_i, "Boarders"] == "YES" or test_home_data_df.loc[person_i, "Tenants"] == "YES":
                Wait3.until(EC.element_to_be_clickable((By.ID, 'houseOccupancy1-label'))).click() # click the "I live here and have boarders or paying guests" button
            else:
                Wait3.until(EC.element_to_be_clickable((By.ID, 'houseOccupancy0-label'))).click() # click this "I live here" button

        elif data["Occupancy"] == "RENTED" or data["Occupancy"] == "LET TO FAMILY" or data["Occupancy"] == "LET TO EMPLOYEES":

            # answer 'Who lives in the house'
            if data["Occupancy"] == "RENTED": # select 'Tenants'
                Wait3.until(EC.element_to_be_clickable((By.ID, "btnhouseRentedTenants-0"))).click()

            elif data["Occupancy"] == "LET TO FAMILY": # select 'Relatives'
                Wait3.until(EC.element_to_be_clickable((By.ID, "btnhouseRentedTenants-1"))).click()

            elif data["Occupancy"] == "LET TO EMPLOYEES": # select 'Employees'
                Wait3.until(EC.element_to_be_clickable((By.ID, "btnhouseRentedTenants-2"))).click()
            else:
                raise Exception("Invalid Input Data Error: Occupancy Column")
            
            # answer "Do you rent your house out as a holiday home or bach?"
            Wait3.until(EC.element_to_be_clickable((By.ID, "btnholidayHomeRented-1"))).click() # always clicked 'No' because the Holiday home occupancy option is handled elsewhere

        elif data["Occupancy"] == "HOLIDAY HOME":
            Wait3.until(EC.element_to_be_clickable((By.ID, 'houseOccupancy3-label'))).click() # click this "This is a holiday home" button

            # answer whether the owner will stay at this holiday home at any point within the next year
            Wait3.until(EC.element_to_be_clickable((By.ID, "btnholidayHomeRented-0"))).click() # always click the 'Yes' button (as we have handled 'No' case during data formatting)

            # click button to say that the owner will stay at the holiday house at some point in the next year (if no then can't get an online quote (handled in data formatting))
            Wait3.until(EC.element_to_be_clickable((By.ID, 'btnholidayHomeRented-0'))).click()


        # select that 'No' significant business is run at the house (the 'Yes' option is handled in data formatting)
        Wait3.until(EC.element_to_be_clickable((By.ID, 'btnhouseUsedForBusiness-1'))).click()


        # enter the owners date of birth
        Wait3.until(EC.visibility_of_element_located((By.ID, "ownerDetails-dob-day"))).send_keys(data["Birthdate_day"])
        Wait3.until(EC.visibility_of_element_located((By.ID, "ownerDetails-dob-month"))).send_keys(data["Birthdate_month"])
        Wait3.until(EC.visibility_of_element_located((By.ID, "ownerDetails-dob-year"))).send_keys(data["Birthdate_year"])


        # answer "In the last three years have you, or any person to be covered by this policy, suffered loss or damage to a house?"
        if data["Type_incident"] !=  "No Incident":
            Wait3.until(EC.element_to_be_clickable((By.ID, "btnpreviousHouseClaims-0"))).click() # click 'Yes' button
            Wait3.until(EC.element_to_be_clickable((By.ID, "lossDamageType-0-toggle"))).click() # click to open the dropdown options for the incident types
            Wait3.until(EC.visibility_of_element_located((By.XPATH, f'//*[@id="lossDamageType-0-menu-options"]/li/div/div[2]/div[text()="{data["Type_incident"]}"]'))).click() # select the correct incident type
            Wait3.until(EC.element_to_be_clickable((By.XPATH, f'//*[@id="lossDamageWhen-0-container"]/div/button/div[2]/div[text()="{data["Incident_year"]}"]'))).click() # selecting the year the incident occured

        else:
            Wait3.until(EC.element_to_be_clickable((By.ID, "btnpreviousHouseClaims-1"))).click() # click 'No' button

        
        # click the "Next: Customise" button to move onto the next page
        Wait3.until(EC.element_to_be_clickable((By.ID, "btnSubmitPage"))).click()


        # wait for the page to load
        time.sleep(5)

        ## selecting the deisred excess level
        # scraping all the excess options from the webpage
        excess_options = WebDriverWait(driver, 20).until(EC.presence_of_all_elements_located((By.XPATH, '//div[@id="excess-container"]/div/button/div[2]/div')))
        excess_values =[funct_defs.convert_money_str_to_int(option.get_attribute("innerHTML")) for option in excess_options]


        # choosing the excess value that is closest to our desired one
        chosen_excess = funct_defs.choose_excess_value(excess_values, int(test_home_data_df.loc[person_i, "Excess"]))

        # selecting the chosen excess value
        chosen_excess_button = Wait10.until(EC.element_to_be_clickable((By.XPATH, f'//div[@id="excess-container"]/div/button/div[2]/div[contains(text(), "{chosen_excess}")]')))
        driver.execute_script("window.scrollTo(0, 0);") # go to the top of the screen

        # attempt to click the chosen excess button (if not on screen scroll down until it is)
        not_clicked = True
        for i in range(100):
            if not_clicked:
                try:
                    chosen_excess_button.click()
                except exceptions.ElementClickInterceptedException:
                    driver.execute_script("window.scrollBy(0, 10);")
                    time.sleep(0.5)
                else:
                    break # if the button is clicked, escape the for loop

        # if the excess was changed, then handle it 
        if chosen_excess != int(test_home_data_df.loc[person_i, "Excess"]):
            # print and output an error message
            print(f"Excess Changed Warning! {test_home_data_df.loc[person_i, "Excess"]} not an option so {chosen_excess} selected", end=" -- ")
            tower_output_df.loc[person_i, "Tower_Error_code"] = f"Excess Changed Warning! {test_home_data_df.loc[person_i, "Excess"]} not an option so {chosen_excess} selected"

            # set the excess for this example to be equal to the newly chosen excess
            tower_output_df.loc[person_i, "Excess"] = chosen_excess


        # click the "Next: " button to move onto the next page (twice)
        for i in range(2):
            next_button = Wait10.until(EC.element_to_be_clickable((By.ID, "btnSubmitPage")))
            driver.execute_script("arguments[0].scrollIntoView();", next_button)
            next_button.click()
            time.sleep(5) # wait for the page to load
        

        # click button to accept terms and conditions
        Wait10.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="questionPrivacyPolicy"]/div[4]/label/div'))).click()


        # select whether the company will be owned by a business or a trust
        if test_home_data_df.loc[person_i, "PolicyHeldByCompany/Trust"].upper() == "YES":
            raise Exception("Invalid Input Data Error: PolicyHeldByCompany/Trust Column-Code not currently equipted to handle 'Yes'")
        elif test_home_data_df.loc[person_i, "PolicyHeldByCompany/Trust"].upper() == "NO":
            Wait3.until(EC.element_to_be_clickable((By.ID, 'btnownedByBusinessOrTrust-1'))).click() # select 'No' button
        else:
            raise Exception("Invalid Input Data Error: Issue with PolicyHeldByCompany/Trust Column")
 

        ## enter a fake name for the person
        # deciding first name (either john or jane)
        if test_home_data_df.loc[person_i, "Gender"].upper() == "MALE": # if person is male enter 'John doe'
            first_name = "John"
        elif test_home_data_df.loc[person_i, "Gender"].upper() == "FEMALE": # if person is male enter 'Jane doe'
            first_name = "Jane"
        else:
            raise Exception("Invalid Input Data Error: Issue with Gender Column")
        
        # typing the names into the input boxes
        Wait3.until(EC.element_to_be_clickable((By.ID, 'txtContactOwner-0-firstName-contactOwner'))).send_keys(first_name)
        driver.find_element(By.ID, 'txtContactOwner-0-lastName-contactOwner').send_keys("Doe")


        # enter fake email address
        Wait3.until(EC.element_to_be_clickable((By.ID, 'txtContactOwner-0-email'))).send_keys(f"{first_name}.Doe@email.com")


        # enter a fake phone number
        Wait3.until(EC.element_to_be_clickable((By.ID, 'txtContactOwner-0-phoneNumbers-0'))).send_keys("021123456")


        # click button to move onto the next page
        Wait3.until(EC.element_to_be_clickable((By.ID, 'btnSubmitPage'))).click()


        # click yes button to answer Yes to 'Do you understand' question: the legal information declaration
        Wait3.until(EC.element_to_be_clickable((By.ID, 'btnlegalDeclaration-0'))).click()


        # click yes button to answer Yes to the 'Do you understand': the important things to call out
        Wait3.until(EC.element_to_be_clickable((By.ID, 'btnexclusions-0'))).click()

        # accept landlord responsibilities
        try:
            # click 'yes' button to accept landlord responsibilities
            Wait3.until(EC.element_to_be_clickable((By.ID, 'btnlandlordDeclaration-0'))).click()
        except:
            pass


        # Select 'Yes' button to say that the property is well maintained
        Wait3.until(EC.element_to_be_clickable((By.ID, 'btnhouseDeclaration-0'))).click()


        # Select 'No' button to say that the person has not had insurance refused within the last 7 years
        Wait3.until(EC.element_to_be_clickable((By.ID, 'btninsuranceHistory-1'))).click()


        # Select 'No' button to say that the person has not had a claim refused within the last 7 years
        Wait3.until(EC.element_to_be_clickable((By.ID, 'btnclaimsDeclined-1'))).click()


        # Select 'No' button to say that the person has not committed a serious crime (Fraud, Arson, Burglary or Theft, Wilful Damage, Sexual Offence, Drugs Conviction) within the last 7 years
        Wait3.until(EC.element_to_be_clickable((By.ID, 'btncriminalHistory-1'))).click()


        # enetering finance details 'Do you owe money on any items you're insuring?'
        if data["Mortgage_bank"] == "No":
            Wait1.until(EC.element_to_be_clickable((By.ID, 'btnmoneyOwed-1'))).click() # click 'No'
        else:
            Wait1.until(EC.element_to_be_clickable((By.ID, 'btnmoneyOwed-0'))).click() # click 'Yes'

            # enter in the name of the finance provider
            Wait1.until(EC.element_to_be_clickable((By.ID, 'financialInterestedParty-0-financialInterestedParty-financial-institution-search'))).send_keys(data["Mortgage_bank"])

            # select the finance provider from the drop down (the first drop down option)
            chosen_financial_provider = Wait3.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="financialInterestedParty-0-financialInterestedParty-financial-institution-search-menu-list"]/li[1]')))
            chosen_financial_provider.click()
            
            # if the financial provider is an 'Other'
            if chosen_financial_provider.text == "Other 1":
                # enter the name of the 'Other' bank
                Wait1.until(EC.element_to_be_clickable((By.ID, 'financialInterestedParty-0-financialInterestedParty-financial-institution-details-description'))).send_keys(data["Mortgage_bank"])

                # enter (fake) financial provider email
                Wait1.until(EC.element_to_be_clickable((By.ID, 'financialInterestedParty-0-financialInterestedParty-financial-institution-details-email'))).send_keys(f"{data["Mortgage_bank"].replace("(", "").replace(")", "")}@email.com")



        

        # click button to move onto the next page
        Wait3.until(EC.element_to_be_clickable((By.ID, 'btnSubmitPage'))).click()
        

        # enter the desired start date of the policy (todays date)
        Wait10.until(EC.element_to_be_clickable((By.ID, 'policyStartDatePicker'))).send_keys(date.today().strftime("%d/%m/%Y"))
        tower_output_df.loc[person_i, "PolicyStartDate"] = date.today().strftime("%d/%m/%Y") # ensuring that the date in the spreadsheet is todays date


        # click button to move onto the next page
        Wait3.until(EC.element_to_be_clickable((By.ID, "btnSubmitPage"))).click()


        ## scraping the home premiums
        # scraping the monthly premium
        monthly_premium = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div/main/div[2]/div/form/fieldset/div[4]/div[2]/button/div[2]/div[2]'))).text # scrape the text string of the monthly premium

        # scraping the yearly premium
        yearly_premium = Wait1.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div/main/div[2]/div/form/fieldset/div[4]/div[3]/button/div[2]/div[2]'))).text # scrape the text string of the monthly premium


        # reformatting the montly and yearly premiums into integers
        monthly_premium, yearly_premium = funct_defs.convert_money_str_to_int(monthly_premium, cents=True), funct_defs. convert_money_str_to_int(yearly_premium, cents=True)

        # returning the monthly/yearly premium and the adjusted agreed value
        return monthly_premium, yearly_premium


    # get time of start of each iteration
    start_time = time.time() 


    # run on the ith car/person
    try:
        # scrapes the insurance premium for a single house/person
        home_premiums = tower_home_scrape_premium(tower_home_data_format(person_i))

        # print the scraping results
        print(home_premiums[0], home_premiums[1], end =" -- ")

        # save the scraped premiums to the output dataset
        tower_output_df.loc[person_i, "Tower_monthly_premium"] = home_premiums[0] # monthly
        tower_output_df.loc[person_i, "Tower_yearly_premium"] = home_premiums[1] # yearly

    except Exception as error_message:
        # convert the error_message into a string
        error_message = str(error_message)

        # defining a list of known error messages
        errors_list = ["Webiste Does Not Quote For This House/ Person", "Invalid Input Data Error"]
        execute_bottom_code = True

        # checking if the error message is one of the known ones
        for error in errors_list:
            # checking if the error message that was returned is a known one
            if  error in error_message:
                print(error_message, end= " -- ")
                tower_output_df.loc[person_i, "Tower_Error_code"] = error_message
                execute_bottom_code = False

        # if the error is not any of the known ones
        if execute_bottom_code:
            # checking if "More info Required pop-up appeared on the screen"
            try: 
                Wait3.until(EC.visibility_of_element_located((By.XPATH, '/html/body/span/span/div/section/div/h5')))
            except exceptions.TimeoutException:
                print("Unknown Error!!", end= " -- ")
                tower_output_df.loc[person_i, "Tower_Error_code"] = error_message
            else:
                if first_attempt == True:
                    print("First attempt failed because of 'More info required!' pop up")
                    person_i = person_i - 1 # if the more info required! issue occured, try again
                    first_attempt = False # set this to False (so that we will not continuosly retry this if the same error keeps occuring)

                    # close and re-open the webdriver
                    driver.close()
                    load_webdriver()
                    return person_i, first_attempt
                else:
                    print("Unknown Error: 'More info required!' pop up appeared", end= " -- ")
                    tower_output_df.loc[person_i, "Tower_Error_code"] = "Unknown Error: 'More info required!' pop up appeared"

            
    

    
    end_time = time.time() # get time of end of each iteration
    print("Elapsed time:", round(end_time - start_time, 2)) # print out the length of time taken
    



### a function to scrape premiums all given examples on tower's website
def tower_auto_scape_all():
    # define how many rows to scrape
    num_cars = len(test_home_data_df)

    # initialising person_i variable
    person_i = 0

    # define a variable that says whether or not this was the first attempt at scraping for this person
    first_attempt = True

    # loop through all examples in test_home_data spreadsheet
    while person_i < num_cars: 

        print(f"{person_i}: TOWER: ", end = "") # print out the iteration number

        # set for this person, the PolicyStartDate to todays date
        test_home_data_df.loc[person_i, "PolicyStartDate"] = datetime.strftime(date.today(), "%d/%m/%Y")

        # run on the ith car/person
        returns = tower_home_premium_scrape(person_i, first_attempt)

        # if something is returned from tower_home_premium_scrape, then save their values
        if returns != None:
            person_i, first_attempt = returns[0], returns[1]
        else:
            first_attempt = True # set to True, if not set to False


        # delete all cookies to reset the page
        try:
            driver.delete_all_cookies()
        except exceptions.TimeoutException: # if we timeout while trying to reset the cookies

                print("\n\nNew Webdriver window\n")
                driver.quit() # quit this current driver
                load_webdriver() # open a new webdriver session

        # iterate up person_i (because we are using a while loop)
        person_i += 1

    funct_defs.export_auto_dataset(tower_output_df, "Tower")




def main():
    # performing all data reading in and preprocessing
    global test_home_data_df, tower_output_df
    test_home_data_df, tower_output_df = funct_defs.dataset_preprocess("Tower")

    # loading the webdriver
    load_webdriver()

    # scrape all of the insurance premiums for the given cars from tower
    tower_auto_scape_all()

    # Close the browser window
    #driver.quit()




main()
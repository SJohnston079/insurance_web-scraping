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

### Defining the key functions for scraping from AA
# defining a function to scrape from the given company website (for an individual person/house)
def aa_home_premium_scrape(person_i):

    # defining a function which take the information from the spreadsheet and formats it so it can be used to scrape premium from aa website
    def aa_home_data_format(person_i):
        # getting the persons birthdate out as a date object (allows us to get the correct format more easily)
        birthdate = test_home_data_df.loc[person_i,'DOB']
        
        # formatting the building type
        building_type = test_home_data_df.loc[person_i, "BuildingType"].upper()
        if building_type in ["APARTMENT", "RETIREMENT UNIT"]:
            raise Exception("Webiste Does Not Quote For This House/ Person, Issue With BuildingType Column")
        elif building_type == "SEMI DETACHED / TOWNHOUSE": 
            building_type = "<7"
        elif "MULTI UNIT" in building_type: # is a multi unit
            number_of_units = int(funct_defs.remove_non_numeric(building_type))
            if (number_of_units < 7 and number_of_units > 0):
                building_type = "<7"
            elif number_of_units < 11:
                building_type = "7-10"
            else:
                raise Exception("Webiste Does Not Quote For This House/ Person, Issue With BuildingType Column")
            
        else:
            building_type = building_type[0] +  building_type[1:].lower() # converts the building type string into a string where the first letter is capitalised and the rest is not


        # formatting the contruction type (The main material that the building was constructed out of)
        construction_type = test_home_data_df.loc[person_i, "ConstructionType"].upper()
        if construction_type in ["MUD BRICK", "NATURAL STONE CLADDING", "ROCKCOTE EPS", "STONEWORK SOLID"]:
            raise Exception("Webiste Does Not Quote For This House/ Person, Issue With ConstructionType Column")
        elif "CONCRETE" in construction_type:
            construction_type = "Concrete"
        elif construction_type == "FIBRE CEMENT CLADDING":
            construction_type = "Fibro"
        elif construction_type == "STUCCO":
            construction_type = "Stucco/Plaster"
        elif construction_type == "TIMBER / WEATHERBOARD":
            construction_type = "Weatherboard/Wood"
        elif construction_type == "METAL CLADDING":
            construction_type = "Aluminium"
        elif construction_type in ["DOUBLE BRICK", "HARDIPLANK/HARDIFLEX", "VINYL CLADDING"]:
            construction_type = construction_type.title() # for all construction materials where we need to capitalise every word
        elif construction_type in ["ALUMINIUM", "BRICK VENEER"]:
            construction_type = construction_type[0] +  construction_type[1:].lower() # for all construction materials where we need to capitalise just the first letter
        else:
            raise Exception("Invalid Input Data Error: ConstructionType Column")

        # formatting the roof type
        roof_type = test_home_data_df.loc[person_i, "RoofType"].upper()
        roof_type = roof_type.replace("PITCHED-", "").replace("FLAT-", "") # remove the prefixes that aren't needed for this website
        if roof_type in ["SHINGLES", "MEMBRANE"]:
            raise Exception("Webiste Does Not Quote For This House/ Person, Issue With RoofType Column")
        elif "TILES" in roof_type: # if roof is made of any type of tiles
            roof_type = "Tiles"
        elif roof_type == "CONCRETE SOLID":
            roof_type = "Concrete"
        elif roof_type in ["ALUMINIUM", "FIBRO", "IRON (CORRUGATED)", "SLATE", "STEEL/COLORBOND", "TIMBER"]:
            roof_type = roof_type.title()
        else:
            raise Exception("Invalid Input Data Error: RoofType Column")
        

        # formatting the number of stories
        num_stories = test_home_data_df.loc[person_i, "NumberOfStories"]
        if num_stories >= 6:
            raise Exception("Webiste Does Not Quote For This House/ Person, Issue With NumberOfStories Column (doesn't give an online quote if the number of stories is greater than or equal to 6)")

        # checking whether there is more than one self contained unit on the property
        self_contained_unit = test_home_data_df.loc[person_i, "NumSelfContainedDwellings"]
        if self_contained_unit > 1:
            raise Exception("Webiste Does Not Quote For This House/ Person, Issue With NumSelfContainedDwellings Column (doesn't give an online quote if the number of self contained dwellings more than 1)")
        elif self_contained_unit <= 0:
            raise Exception("Invalid Input Data Error: NumSelfContainedDwellings, House cannot have < 1 dwelling")

        # formatting the number of covered external car spaces (excluding internal garage)
        covered_car_spaces = test_home_data_df.loc[person_i, "CoveredExternalCarSpaces"] 
        if int(covered_car_spaces) >= 4: # any number greater than or equal to 4 is mapped to 4+
            covered_car_spaces = "4+"

        ## formatting a variable to allow us to input all of the outbuildings we wish to input
        outbuildings_presence_encoding = "" # initialising this variable

        # iterate through all of the outbuilding types defined below and encode whether they 
        outbuilding_types = ["Balcony", "SwimmingPool", "SportCourt", "GardenShed", "Shed", "WaterTanks"]
        for i in range(len(outbuilding_types)):
            outbuilding = test_home_data_df.loc[person_i, outbuilding_types[i]].upper()
            if outbuilding == "YES":
                outbuildings_presence_encoding += str(i+1) # add a number to the encoding to represent the given type of outbuilding only if it is present
            elif outbuilding != "NO": # checking that the data in the given column is not invalid
                raise Exception(f"Invalid Input Data Error: {outbuilding_types[i]} column")
        

        ## formatting a bunch of variables relevant to the house being insured and the person getting the policy
        # Do you run any type of business from your house?
        business_use = test_home_data_df.loc[person_i, "BusinessUser"]
        if business_use in ("Yes-Other", "Yes->50% of house used for business"):
            raise Exception("Webiste Does Not Quote For This House/ Person, Issue With BusinessUser Column")
        elif business_use in ("Yes-Home office", "Yes-Hobby farm", "Yes-Bed and breakfast", "No"):
            business_use = business_use.replace("Yes-", "")
        else: # if the input is not recognised
            raise Exception(f"Invalid Input Data Error: BusinessUser column")

        # Is you house watertight, structurally sound, secure and well maintained?
        well_maintained = test_home_data_df.loc[person_i, "HouseWellMaintained"].upper()
        if well_maintained == "NO":
            raise Exception("Webiste Does Not Quote For This House/ Person, Issue With HouseWellMaintained Column (Has HouseWellMaintained='No')")
        elif well_maintained != "YES":
            raise Exception(f"Invalid Input Data Error: HouseWellMaintained column")
        
        # Have Heritage New Zealand or your local council placed any restrictions or preservation orders on this house? 
        heritage = test_home_data_df.loc[person_i, "HeritageProperty"].upper()
        if heritage == "YES-HERITAGE" or heritage == "YES-LOCAL_COUNCIL":
            heritage = heritage.replace("YES-", "").title()
        elif heritage == "NO":
            heritage = heritage.title()
        else:
            raise Exception(f"Invalid Input Data Error: HeritageProperty column")
        
        # In the last seven years have you or anyone covered by this policy had any insurance refused, cancelled, special terms imposed, renewal not offered or a claim declined?
        insurance_refused_last7years = test_home_data_df.loc[person_i, "Insurance Refused In Last 7 Years"]
        if insurance_refused_last7years in ("Yes-Fraud", "Yes-Non-disclosure", "Yes-Misrepresentation", "Yes-Breach of policy conditions"): # checking if the person is not insurable due to insurance_refused_last7years
            raise Exception("Webiste Does Not Quote For This House/ Person, Issue With Insurance Refused In Last 7 Years Column")
        elif insurance_refused_last7years in ("No", "Yes-Other"): # checking whether it is a valid value
            insurance_refused_last7years = insurance_refused_last7years.title()
        else: # if the input is not valid
            raise Exception(f"Invalid Input Data Error: Insurance Refused In Last 7 Years column")

        # In the last seven years have you or anyone covered by this policy had any criminal convictions?
        crime_last7years = test_home_data_df.loc[person_i, "Crime in Last 7 Years"]
        if crime_last7years in ("Yes-Fraud", "Yes-Arson", "Yes-Burglary", "Yes-Theft", "Yes-Drugs"):
            raise Exception("Webiste Does Not Quote For This House/ Person, Issue With Crime in Last 7 Years Column")
        elif crime_last7years in  ("No", "Yes-Other"):  # checking whether it is a valid value
            crime_last7years = crime_last7years.title()
        else: # if the input is not valid
            raise Exception(f"Invalid Input Data Error: HouseWellMaintained column")
        
        ## formatting incidents
        # getting the date of the incident
        incident_date = test_home_data_df.loc[person_i, "Date_of_incident"]
        incident_year = incident_date.strftime("%Y") # getting the year of the incident

        # if the incident occured within the last 3 years
        if funct_defs.check_date_range(incident_date, 3):
            type_incident = test_home_data_df.loc[person_i, "Type_incident"]
        else: # is the incident did not occur within the last 3 years then it doesn't matter
            type_incident = "No Incident"



        ## define a dict to store information for a given house/person
        data  = {'Has_permanent_residents': not (test_home_data_df.loc[person_i,'ShortTermTenancy'] == 'Yes'), # if it is not a short term enancy then it must have permanent residents
                'Street_address':f"{test_home_data_df.loc[person_i,'Street_number']} {test_home_data_df.loc[person_i,'Street_name']} {test_home_data_df.loc[person_i,'Street_type']}",
                "Birthdate_day":str(int(birthdate.strftime("%d"))),
                "Birthdate_month":birthdate.strftime("%B"),
                "Birthdate_year":birthdate.strftime("%Y"),
                "Other_policies": "YES" in [test_home_data_df.loc[person_i, insurance_type].upper() for insurance_type in ("CarInsurance", "ContentsInsurance", "FarmInsurance", "BoatInsurance")],
                "Building_type":building_type,
                "Construction_type":construction_type,
                "Roof_type":roof_type,
                "Num_stories":num_stories,
                "Covered_car_spaces":covered_car_spaces,
                "Outbuildings_presence_encoding":outbuildings_presence_encoding,
                "Business_use":business_use,
                "Heritage":heritage,
                "Insurance_refused_last7years":insurance_refused_last7years,
                "Crime_last7years":crime_last7years,
                "Incident_year":incident_year,
                "Type_incident":type_incident
                 }
        
        return data


    # scrapes the insurance premium for a single vehicle/person at aa
    def aa_home_scrape_premium(data):
        
        ## opening a window to get home insurance premium from
        if test_home_data_df.loc[person_i,'Occupancy'].upper() == 'UNOCCUPIED': # if the house is unoccupied
            # if the house is unoccupied, we must call in to get cover (so cannot be scraped)
            raise Exception("Webiste Does Not Quote For This House/ Person, Issue With Occupancy Column")
        
        elif test_home_data_df.loc[person_i,'Occupancy'].upper() in ['OWNER OCCUPIED', 'HOLIDAY HOME']: # if can be covered by standard home insurance
            # selects 'Home Insurance'
            home_insurance_button = Wait10.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="page"]/main/div[1]/div[2]/div[2]/div/div[1]/div[1]/div/ul/li[1]/a/div[2]/span'))) 
            driver.execute_script("arguments[0].scrollIntoView();", home_insurance_button)
            time.sleep(1) # wait for the page to load
            home_insurance_button.click() 

        elif test_home_data_df.loc[person_i,'Occupancy'].upper() in ['RENTED', 'LET TO FAMILY', 'LET TO EMPLOYEES']:
            # selects 'Landlord Insurance'
            landlord_insurance_button = Wait10.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="page"]/main/div[1]/div[2]/div[2]/div/div[1]/div[1]/div/ul/li[6]/a/div[2]/span')))
            driver.execute_script("arguments[0].scrollIntoView();", landlord_insurance_button)
            time.sleep(1) # wait for the page to load
            landlord_insurance_button.click() 
        
        else:
            raise Exception("Invalid Input Data Error: Occupancy Column")


        time.sleep(1) # wait for the page to load
        driver.switch_to.window(driver.window_handles[1]) # Switch back to the second tab (that we have just opened)



        # select whether the person is an AA member
        if test_home_data_df.loc[person_i, 'AAMember'].upper() == "YES":
            driver.find_element(By.XPATH, '//*[@id="aaMembershipDetailButtons"]/label[1]').click()
        else:
            driver.find_element(By.XPATH, '//*[@id="aaMembershipDetailButtons"]/label[2]').click()

        # selects whether or not the house has a permanent resident
        if data["Has_permanent_residents"]: # if house does have permanent resident(s)
            driver.find_element(By.XPATH, '//*[@id="residencyTypeButtons"]/label[1]/span/span').click()
        else: # if house does NOT have permanent resident(s) (e.g. is a holiday house)
            driver.find_element(By.XPATH, '//*[@id="residencyTypeButtons"]/label[2]/span/span').click()


        # Select the cover type (Am currently just always selecting House only insurance)
        Wait10.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="coverTypeButtons"]/label[1]/span'))).click()
        time.sleep(2) # wait for the page to load


        ## entering the house address
        # entering suburb + poscode + city
        suburb_input_box = driver.find_element(By.ID, "address.suburbPostcodeRegionCity")
        suburb_input_box.send_keys(str(test_home_data_df.loc[person_i,'Postcode'])) # entering the postcode into the input box
        suburb_option = Wait3.until(EC.element_to_be_clickable((By.XPATH, f'//*[@id="insuranceOptions"]/fieldset[4]/div[1]/div/ul/li[contains(text(), "{test_home_data_df.loc[person_i,'Suburb'].strip()}")]')))
        driver.execute_script("arguments[0].scrollIntoView();", suburb_option) # scroll so that we can see the option we want to click
        suburb_option.click() # select the dropdown with the correct suburb

        # entering the street address (street name, type and number)
        street_address_input_box = driver.find_element(By.ID, "address.streetAddress")
        street_address_input_box.send_keys(data["Street_address"]) # entering the postcode into the input box


        # checking if there are any issue with the chosen address
        try:
            driver.find_element(By.XPATH, '//*[@id="insuranceOptions"]/fieldset[5]/div/label').click() # clicking a random part of the screen to exit the input box
            Wait3.until(EC.visibility_of_element_located((By.CLASS_NAME, 'jeopardy-group')))

        except exceptions.TimeoutException:
            pass
        else:
            raise Exception("Webiste Does Not Quote For This House/ Person: Issue With 'The Address of the House' (Want you to call in to handle details)")


        # scraping the resulting address
        aa_output_df.loc[person_i, "AA_selected_address"] = f"{street_address_input_box.get_attribute("value")} {suburb_input_box.get_attribute("value")}"


        # entering the policy holders data of birth
        Select(driver.find_element(By.ID, 'dateOfBirth-day')).select_by_value(data["Birthdate_day"])
        Select(driver.find_element(By.ID, 'dateOfBirth-month')).select_by_visible_text(data["Birthdate_month"])
        Select(driver.find_element(By.ID, 'dateOfBirth-year')).select_by_visible_text(data["Birthdate_year"])


        # if the person has other policies with this insurer (AA)
        if data["Other_policies"]:
            driver.find_element(By.XPATH, '//*[@id="existingSuncorpPoliciesButtons"]/label[1]/span/span').click() # click the 'Yes' button
        else:
            driver.find_element(By.XPATH, '//*[@id="existingSuncorpPoliciesButtons"]/label[2]/span/span').click() # click the 'No' button


        # select who the persons most recent insurer was
        recent_insurer_dropdown = Select(driver.find_element(By.ID, "previousInsurer"))
        recent_insurer_dropdown.options[1] # choose the second optgroup 'All Insurers'
        recent_insurer_dropdown.select_by_visible_text(test_home_data_df.loc[person_i, "CurrentInsurer"]) # select the correct current/previous insurer
        

        # click the 'continue' button to move onto the next page
        Wait10.until(EC.element_to_be_clickable((By.ID, '_eventId_submit'))).click()
        
        # checks if the error message bar that says "We couldn't recognise your address. Please re-enter it again below" appears. Handle it if it does
        try:
            Wait3.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="*.errors"]/ul/li')))

        except: # if it doesn't appear, continue
            pass
        else: # if the error message bar does appear then handle it
            all_options = Wait3.until(EC.visibility_of_all_elements_located((By.XPATH, '/html/body/div[2]/main/div/div[2]/form/fieldset[4]/div[4]/div[2]/div/label/span'))) # find all potential options

            # initialising these variables
            best_match_option = ""
            best_match_score = 0
            actual_address = f"{data["Street_address"].title()}, {test_home_data_df.loc[person_i, "Suburb"].title()}, {test_home_data_df.loc[person_i, "Postcode"]}"

            # going through all options and selecting the best one
            for option in all_options:
                score = fuzz.partial_ratio(actual_address, option.text) # calculates how similar the given option is to the actual address
                if score > best_match_score:
                    best_match_score = score
                    best_match_option = option
            
            # select the best option
            best_match_option.click()

            # save the selected address
            aa_output_df.loc[person_i, f"AA_selected_address"] = best_match_option.text

            # click the 'continue' button to move onto the next page
            Wait10.until(EC.element_to_be_clickable((By.ID, '_eventId_submit'))).click()
            

        # inputting the building type
        if data["Building_type"] == "<7":
            Wait10.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="buildingTypeButtons"]/label[3]/span/span/small'))).click()
        elif data["Building_type"] == "7-10":
            Wait10.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="buildingTypeButtons"]/label[4]/span/span'))).click()
        else:   
            Wait10.until(EC.element_to_be_clickable((By.XPATH, f'//*[@id="buildingTypeButtons"]/label/span/span[contains(text(), "{data["Building_type"]}")]'))).click()


        # select whether or not the house is a part of a body corporate
        if test_home_data_df.loc[person_i, "BodyCorporate"].upper() == "YES": 
            driver.find_element(By.XPATH, '//*[@id="strataTitleButtons"]/label[1]/span/span').click()
        elif test_home_data_df.loc[person_i, "BodyCorporate"].upper() == "NO":
            driver.find_element(By.XPATH, '//*[@id="strataTitleButtons"]/label[2]/span/span').click()
        else:
            raise Exception("Invalid Input Data Error: BodyCorporate Column")
        

        ## inputting the materials that the house is made of
        # inputting the construction material (the main material the house is made of)
        driver.find_element(By.XPATH, f'//*[@id="externalWallMaterialButtons"]/label/span/span[contains(text(), "{data["Construction_type"]}")]').click()

        # selecting the roof type (the main material the roof of the house is made of)

        Wait3.until(EC.element_to_be_clickable((By.XPATH, f'//*[@id="roofMaterialButtons"]/label/span/span[contains(text(),"{data["Roof_type"]}")]'))).click()


        ## other house details
        # what year was the house built
        driver.find_element(By.ID, "constructionYear").send_keys(str(test_home_data_df.loc[person_i, "YearBuilt"]))

        # house security (is there an alarm, and is it monitored by a security provider)
        if test_home_data_df.loc[person_i, "HouseSecurity"].upper() != "NO SECURITY":
            driver.find_element(By.XPATH, '//*[@id="homeDetails"]/fieldset[4]/h2').click() # clicking to exit the input box 

            # selecting the correct option
            if "NOT MONITORED" in test_home_data_df.loc[person_i, "HouseSecurity"].upper(): # Burglar alarm, not monitored by a security provider
                Wait3.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="alarms-container"]/div/div/label[1]/span'))).click()
            else: # Burglar alarm, monitored by a security provider
                Wait3.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="alarms-container"]/div/div/label[2]/span'))).click()

        # click 'Continue' button to move onto the next page
        driver.find_element(By.ID, "_eventId_submit").click()
        
        # input how many levels the house has
        Wait10.until(EC.element_to_be_clickable((By.XPATH, f'//*[@id="buildingCover"]/fieldset[1]/div/div/div/div/label/span[contains(text(), "{data["Num_stories"]}")]'))).click()

        # input whether the house has more than 1 self contained dwelling (It must have exactly 1 self contained dwelling, as any other value would have raised an error inside of the aa_home_data_format function)
        Wait10.until(EC.element_to_be_clickable((By.XPATH, f'//*[@id="noOfDwellingUnitsButtons"]/label[2]/span/span'))).click()

        # selects the number of covered car spaces in the house
        driver.find_element(By.XPATH, f'//*[@id="carSpaceButtons"]/label/span/span[contains(text(), "{data["Covered_car_spaces"]}")]').click()

        # inputting details about outbuildings
        for encoded_val in data["Outbuildings_presence_encoding"]: # iterate through the encoding and click the buttons to indicate that the property has the given outbuildings
            driver.find_element(By.XPATH, f'//*[@id="outbuildings-container"]/div/div/label[{encoded_val}]/span').click()

        # inputting the size of the house (in square metres)
        driver.find_element(By.ID, "buildingArea").send_keys(str(test_home_data_df.loc[person_i, "DwellingFloorArea"]))

        # handling SumInsured
        # clicking button to 'estimate replacement cost' (the SumInsured)
        Wait3.until(EC.element_to_be_clickable((By.ID, "estimateReplacementCost"))).click()
        
        # check if the error page appears (need more information)
        try:
            Wait3.until(EC.presence_of_element_located((By.ID, 'jeopardy-aainz')))
            
        except exceptions.TimeoutException: # if the error page does not appear
            pass
        else:
            # checking what type of error occured
            text = driver.find_element(By.XPATH, '//*[@id="command"]/fieldset/div[1]/ul/li').get_attribute("innerHTML").strip()
            if text == "the size of your house":
                raise Exception("Webiste Does Not Quote For This House/ Person: Issue With DwellingFloorArea Column (Want you to call in): (Value is likely too high)")
            else:
                raise Exception("Unknown Error")

        # scraping the estimated replacement cost
        aa_output_df.loc[person_i, "AA_estimated_replacement_cost"] = funct_defs.convert_money_str_to_int(Wait3.until(EC.presence_of_element_located((By.ID, "estRebuildCost"))).text)

        # inputting the desired SumInsured
        driver.find_element(By.ID, "buildingSum").send_keys(str(test_home_data_df.loc[person_i, "SumInsured"]))


        # click button to go to the next page
        driver.find_element(By.ID, "_eventId_submit").click()
        time.sleep(1) # wait for the page to load

        try:
            # check if the error ('jeopardy') page appears. If it doesn't, then continue
            Wait3.until(EC.presence_of_element_located((By.ID, "jeopardy-aainz")))

        except: # if we are not on the error page, then we can continue as normal (the SumInsured value is valid)
            pass
        else:
            # if the sum insured value is larger than the AA_estimated_cost, then it must be too large (as we can only get to this part if we are on the error page (we check it above))
            if test_home_data_df.loc[person_i, "SumInsured"] > aa_output_df.loc[person_i, "AA_estimated_replacement_cost"]:
                raise Exception("Webiste Does Not Quote For This House/ Person: SumInsured Column: Value Too Large")
            else: # if the sum insured value is smaller than the AA_estimated_cost, then it must be too small (as we can only get to this part if we are on the error page (we check it above))
                raise Exception("Webiste Does Not Quote For This House/ Person: SumInsured Column: Value Too Small")
        
        
        ## selecting the excess on the policy
        # scraping all the excess options from the webpage
        excess_options = Wait10.until(EC.presence_of_all_elements_located((By.XPATH, '//*[@id="buildingExcessContainer"]/label/span[1]')))
        excess_values = [funct_defs.convert_money_str_to_int(option.text) for option in excess_options]

        # choosing the excess value that is closest to our desired one
        chosen_excess = funct_defs.choose_excess_value(excess_values, int(test_home_data_df.loc[person_i, "Excess"]))

        # selecting (clicking) the button with the correct excess
        driver.find_element(By.XPATH, f'//*[@id="buildingExcessContainer"]/label/span[contains(text(), "{chosen_excess}")]').click()

        # if the excess was changed, then handle it 
        if chosen_excess != int(test_home_data_df.loc[person_i, "Excess"]):
            # print and output an error message
            print(f"Excess Changed Warning! {test_home_data_df.loc[person_i, "Excess"]} not an option so {chosen_excess} selected", end=" -- ")
            aa_output_df.loc[person_i, "AMI_Error_code"] = f"Excess Changed Warning! {test_home_data_df.loc[person_i, "Excess"]} not an option so {chosen_excess} selected"

            # set the excess for this example to be equal to the newly chosen excess
            aa_output_df.loc[person_i, "Excess"] = chosen_excess


        # select if you want to have to have glass excess
        if test_home_data_df.loc[person_i, "GlassExcess"].upper() == "YES":
            driver.find_element(By.XPATH, '//*[@id="quoteSchedule"]/div[1]/div/div[1]/fieldset[5]/div/label/span').click()
        elif test_home_data_df.loc[person_i, "GlassExcess"].upper() != "NO":
            raise Exception("Invalid Input Data Error: GlassExcess Column")
        

        time.sleep(2) # wait for the page to load

        # scrape the yearly premium string
        yearly_premium = driver.find_element(By.XPATH, '//*[@id="tab-contentToggle__Annually"]/div/div/span[1]/h1').get_attribute('innerHTML')

        # scrape the monthly premium string
        Wait10.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/main/div/form/div[1]/div/div[2]/div[2]/div/div[1]/div/div[1]/div[1]/button[2]'))).click()
        time.sleep(5) # wait for the page to load
        monthly_premium = driver.find_element(By.XPATH, '//*[@id="tab-contentToggle__Monthly"]/div/div/span[1]/h1').get_attribute('innerHTML')         
        
        # reformatting the montly and yearly premiums into integers
        monthly_premium, yearly_premium = funct_defs.convert_money_str_to_int(monthly_premium, cents=True), funct_defs. convert_money_str_to_int(yearly_premium, cents=True)

        # returning the monthly/yearly premium and the adjusted agreed value
        return monthly_premium, yearly_premium


    # get time of start of each iteration
    start_time = time.time() 


    # run on the ith car/person
    try:
        # scrapes the insurance premium for a single house/person
        home_premiums = aa_home_scrape_premium(aa_home_data_format(person_i))

        # print the scraping results
        print(home_premiums[0], home_premiums[1], end =" -- ")

        # save the scraped premiums to the output dataset
        aa_output_df.loc[person_i, "AA_monthly_premium"] = home_premiums[0] # monthly
        aa_output_df.loc[person_i, "AA_yearly_premium"] = home_premiums[1] # yearly

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
                aa_output_df.loc[person_i, "AA_Error_code"] = error_message
                execute_bottom_code = False

        # if the error is not any of the known ones
        if execute_bottom_code:
            print("Unknown Error!!", end= " -- ")
            aa_output_df.loc[person_i, "AA_Error_code"] = error_message
    

    
    # if there is more than one window open
    if len(driver.window_handles) > 1:
        # close the window we have just been scraping from
        driver.close()
        driver.switch_to.window(driver.window_handles[0]) # Switch back to the first tab (as the current tab we are on no longer exists)


    end_time = time.time() # get time of end of each iteration
    print("Elapsed time:", round(end_time - start_time, 2)) # print out the length of time taken
    


### a function to scrape premiums all given examples on aa's website
def aa_auto_scape_all():
    # defining the number of rows to scrape as the number of rows in the input excel spreadsheet
    num_rows_to_scrape = len(test_home_data_df)

    # Open the webpage AA webpage
    driver.get("https://www.aainsurance.co.nz/")
    time.sleep(1) # wait for the page to load
    Wait10.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="page"]/main/div[1]/div[2]/div[2]/div/div[1]/div[1]/div/div/div/div[1]'))).click() # click the 'Get a quote button'

    # loop through all examples in test_home_data spreadsheet
    for person_i in range(0, num_rows_to_scrape): 

        print(f"{person_i}: AA: ", end = "") # print out the iteration number

        # set for this person, the PolicyStartDate to todays date
        test_home_data_df.loc[person_i, "PolicyStartDate"] = datetime.strftime(date.today(), "%d/%m/%Y")

        # run on the ith car/person
        aa_home_premium_scrape(person_i)
        

        # delete all cookies to reset the page
        try:
            driver.delete_all_cookies()
        except exceptions.TimeoutException: # if we timeout while trying to reset the cookies

                print("\n\nNew Webdriver window\n")
                driver.quit() # quit this current driver
                load_webdriver() # open a new webdriver session


    funct_defs.export_auto_dataset(aa_output_df, "AA")




def main():
    # performing all data reading in and preprocessing
    global test_home_data_df, aa_output_df
    test_home_data_df, aa_output_df = funct_defs.dataset_preprocess("AA")

    # loading the webdriver
    load_webdriver()

    # scrape all of the insurance premiums for the given cars from aa
    aa_auto_scape_all()

    # Close the browser window
    driver.quit()

main()
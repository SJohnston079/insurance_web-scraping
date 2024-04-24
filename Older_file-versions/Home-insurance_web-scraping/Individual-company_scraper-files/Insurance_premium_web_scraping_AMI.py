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

        # formatting the address of the house
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
        def send_keys_data_entry(id, column_name):
            # enters the values into the input box
            driver.find_element(By.ID, id).send_keys(str(test_home_data_df.loc[person_i, column_name]))

            # click out of the input box
            driver.find_element(By.ID, "hasGarageContainer").click() 
            time.sleep(1) # wait for the page to load

            # check if the error message saying "Sorry, we're unable to complete your quote online as we require further information about your house" appears
            try:
                Wait1.until(EC.visibility_of_element_located((By.ID, f"{id}StopperError"))) 
                print(f"Invalid Input Data Error: {column_name} Column", end=" -- ")
                return f"Invalid Input Data Error: {column_name} Column"
            except exceptions.TimeoutException:
                if column_name == "SumInsured":
                    try:
                        Wait1.until(EC.visibility_of_element_located((By.ID, "bldgSumInsuredUnderStopperError")))
                        print(f"Invalid Input Data Error: SumInsured Column: Value Too Small", end=" -- ")
                        return f"Invalid Input Data Error: SumInsured Column: Value Too Small"
                    except exceptions.TimeoutException:
                        return None

        # checking that the returned value from ami_home_data_format is not a string
        if isinstance(data, str):
            return data # if the data variables is a string, then it must be an error message, so we return it

        # Opens the AMI House Insurance webpage
        driver.get("https://www.ami.co.nz/house-insurance")

        
        ## selecting the correct type of house insurance
        if data["Occupancy"] in ['UNOCCUPIED', 'BOARDING HOUSE']: # if the house is unoccupied or is a boarding house
            # if the house is unoccupied or is a boarding house, we must call in to get cover (so cannot be scraped)
            return "Webiste Does Not Quote For This House/ Person, Issue With Occupancy Column" 
        
        elif data["Occupancy"] in ['OWNER OCCUPIED', 'HOLIDAY HOME']: # if can be covered by standard home insurance
            # selects 'Premier House' (clicks 'Get a quote button')
            Wait10.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="rootnode"]/div[2]/div/div[2]/div[2]/div/div/div[1]/div/div/a[1]'))).click() 

        elif data["Occupancy"] in ['RENTED', 'LET TO FAMILY/EMPLOYEES']:
            # selects 'Premier Rental Property (Landlord Insurance)' (clicks 'Get a quote button')
            Wait10.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="rootnode"]/div[2]/div/div[2]/div[2]/div/div/div[3]/div/div/a[1]'))).click() 
        else:
            print(f"Data Entry ERROR - Occupancy {test_home_data_df.loc[person_i,'Occupancy']} not understood", end=' -- ')
            return "Invalid Input Data Error: Occupancy Column"

        

        # select whether there more than one self-contained dwelling in the house
        if int(test_home_data_df.loc[person_i, "NumSelfContainedDwellings"]) > 1:
            print("Webiste Does Not Quote For This House/ Person, Issue With Occupancy Column", end=" -- ")
            return "Webiste Does Not Quote For This House/ Person, Issue With Occupancy Column"
        else:
            Wait3.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="initialQuestions"]/div[2]/div[1]/fieldset/div[2]/ul/li[2]/label'))).click() # click 'No' Button



        # states whether any part of the property is currently under renovations
        if data["Current_renovations"]: # if house IS currently undegoing renovations
            driver.find_element(By.XPATH, '//*[@id="initialQuestions"]/div[2]/div[2]/fieldset/div[2]/ul/li[1]/label').click() # click 'Yes' Button
            Select(driver.find_element(By.ID, "propertyRenovationType")).select_by_value("nonstructuralupto75000") # select the options saying that the renovations are non-structural worth less than 75K, (As all other options are filtered out as we cannot get quote online for them)

        else: # if house IS NOT currently undegoing renovations
            driver.find_element(By.XPATH, '//*[@id="initialQuestions"]/div[2]/div[2]/fieldset/div[2]/ul/li[2]/label').click() # click 'No' Button
        

        # click 'Continue' button
        driver.find_element(By.ID, "continueToYourDetailsPage").click()


        ## entering the house address
        # entering suburb + postcode + city
        Wait3.until(EC.element_to_be_clickable((By.ID, "addressFinder"))).send_keys(data["Street_address"]) # entering the street address into the input box
        time.sleep(3) # wait for the page to load
        address_options = driver.find_elements(By.XPATH, "//*[@id='ui-id-2']/li/a")

        # Finds the option with the highest score. A higher score means the option is more similar to the data from the database (as represented by the Full_address string)
        best_match_option = max(address_options, key=lambda option: fuzz.ratio(data["Full_address"], option.text))
        output_df.loc[person_i, "AMI_selected_address"] = best_match_option.text # outputting the selected address
        best_match_option.click() # selecting the highest score option


        # stating whether or not the house is used for business use
        if test_home_data_df.loc[person_i, "BusinessUser"].upper() == "NO":
            try:
                # select the option that says that house is only for private/residential use
                Select(driver.find_element(By.ID, "usage")).select_by_value("ResidentialPrivateUseOnly") 

            except exceptions.NoSuchElementException: # if unable to select the correct dropdown option, check if the dropdown has been replaced with 'Yes'/'No' buttons
                # Click the 'Yes' button, answering the question "Is the house used solely for residential purposes?"
                driver.find_element(By.XPATH, '//*[@id="yourHomeOverview"]/div[2]/div[2]/fieldset/div[2]/ul/li[1]/label').click() 
        else: 
            return "Invalid Input Data Error: BusinessUser Column"

        # waiting for the page to load
        time.sleep(1)

        ## entering in the house occupancy details
        # finding the select box to enter the occupancy details
        occupancy_dropdown = Select(driver.find_element(By.ID, 'occupationType'))

        ## selecting the correct type of house insurance
        if data["Occupancy"] == 'OWNER OCCUPIED':
            if test_home_data_df.loc[person_i, "Tenants"].title() == "Yes":
                occupancy_dropdown.select_by_value("OwnerAndTenants") # select "Owner+Tenants" option
            elif test_home_data_df.loc[person_i, "Boarders"].title() == "Yes":
                occupancy_dropdown.select_by_value("OwnerAndBoarder") # select "Owner+Boarders" option
            else:
                occupancy_dropdown.select_by_value("OwnerOccupied") # select "Owner Occupied" option

        elif data["Occupancy"] == 'HOLIDAY HOME':
            # select "Holiday Home" option
            occupancy_dropdown.select_by_value("HolidayHome") 
        elif data["Occupancy"] == 'RENTED':
            # select "Let to Tenants" option
            occupancy_dropdown.select_by_value("LetToTenants")
        elif data["Occupancy"] == 'LET TO FAMILY/EMPLOYEES':
            # select "Employee/Relative" option
            occupancy_dropdown.select_by_value("LetToEmployeeOrRelative")
        else:
            print(f"Data Entry ERROR - Occupancy {test_home_data_df.loc[person_i,'Occupancy']} Data/Format not understood", end=' -- ')
            return "Invalid Input Data Error: Occupancy Column"

        # handling "Does more than one tenancy operate at the property?" question if it is present
        try:
            # checking if the button to ask if there is more than one tenancy is present on the property
            many_tenancy_No_button = Wait1.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="ownersOccupants"]/div[2]/div[2]/fieldset/div[2]/ul/li[2]/label')))

            if test_home_data_df.loc[person_i, "NumTenanciesOnProperty"] <= 1: # if there are 1 or less tenancies on the property
                many_tenancy_No_button.click() # click 'No' button
            else:
                return "Webiste Does Not Quote For This House/ Person, Issue With NumTenanciesOnProperty Column" # cannot get an online quote for an example with more than 1 tenancy
            
        except exceptions.TimeoutException: # if the button is not present, then we dont need to handle it
            pass 
        
        # handling "Is the tenancy agreement short-term of 30 days or less?" question if it is present
        try:
            # checking if the button to ask if there is more than one tenancy is present on the property
            many_tenancy_No_button = Wait1.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="ownersOccupants"]/div[2]/div[3]/fieldset/div[2]/ul/li[2]/label')))

            if test_home_data_df.loc[person_i, "ShortTermTenancy"] == "No": # if there is no 'short term tenancy' at the house
                many_tenancy_No_button.click() # click 'No' button
            elif test_home_data_df.loc[person_i, "ShortTermTenancy"] == "Yes":
                return "Webiste Does Not Quote For This House/ Person, Issue With ShortTermTenancy Column" # cannot get an online quote for an example where there is a short term tenancy
            else:
                return "Invalid Input Data Error: ShortTermTenancy Column"
            
        except exceptions.TimeoutException: # if the button is not present, then we dont need to handle it
            pass 
        
        

        ## entering the owners birthdate
        driver.find_element(By.ID, "ownerDOB-full_0").send_keys(data['Birthdate'])


        ## entering the home owners gender
        if test_home_data_df.loc[person_i, 'Gender'].upper() == "MALE": 
            driver.find_element(By.XPATH, '//*[@id="rbIncidents0"]/li[1]').click()
        else: # is female
            driver.find_element(By.XPATH, '//*[@id="rbIncidents0"]/li[2]').click()
        

        ## entering rental details
        # if this is premier rental insurance
        if data["Occupancy"] in ['RENTED', 'LET TO FAMILY/EMPLOYEES']:

            # do you want lost rent cover (for if the house is damaged and cannot be lived in)
            if test_home_data_df.loc[person_i, 'LostRentCover'].upper() == "YES": 
                driver.find_element(By.XPATH, '//*[@id="occupantsAddons"]/div[2]/div[1]/fieldset/div[2]/ul/li[1]/label').click()
            elif test_home_data_df.loc[person_i, 'LostRentCover'].upper() == "NO":
                driver.find_element(By.XPATH, '//*[@id="occupantsAddons"]/div[2]/div[1]/fieldset/div[2]/ul/li[2]/label').click()
            else:
                return "Invalid Input Data Error: LostRentCover Column"
            
            # do you want to cover for unexpected vacancy (if your tenant leaves with giving proper notice)
            if test_home_data_df.loc[person_i, 'UnexpectedVacancyCover'].upper() == "YES": 
                driver.find_element(By.XPATH, '//*[@id="occupantsAddons"]/div[2]/div[2]/div[2]/ul/li[1]/label').click()
            elif test_home_data_df.loc[person_i, 'LostRentCover'].upper() == "NO":
                driver.find_element(By.XPATH, '//*[@id="occupantsAddons"]/div[2]/div[2]/div[2]/ul/li[2]/label').click()
            else:
                return "Invalid Input Data Error: UnexpectedVacancyCover Column"
            
            # do you want to cover for theft or damage by tenants
            if test_home_data_df.loc[person_i, 'TheftOrDamageByTenantsCover'].upper() == "YES": 
                driver.find_element(By.XPATH, '//*[@id="occupantsAddons"]/div[2]/div[3]/fieldset/div[2]/ul/li[1]/label').click()
            elif test_home_data_df.loc[person_i, 'LostRentCover'].upper() == "NO":
                driver.find_element(By.XPATH, '//*[@id="occupantsAddons"]/div[2]/div[3]/fieldset/div[2]/ul/li[2]/label').click()
            else:
                return "Invalid Input Data Error: TheftOrDamageByTenantsCover Column"
            
            # do you want to cover for chattels (everything in the rental property that isn't a part of the house (e.g. furniture, appliances, cutlery, ...))
            if test_home_data_df.loc[person_i, 'ChattelsCover'].upper() == "YES": 
                driver.find_element(By.XPATH, '//*[@id="occupantsAddons"]/div[3]/div[1]/fieldset/div[2]/ul/li[1]/label').click()
            elif test_home_data_df.loc[person_i, 'LostRentCover'].upper() == "NO":
                driver.find_element(By.XPATH, '//*[@id="occupantsAddons"]/div[3]/div[1]/fieldset/div[2]/ul/li[2]/label').click()
            else:
                return "Invalid Input Data Error: ChattelsCover Column"


        # click the 'Continue' button to move onto the next page
        driver.find_element(By.ID, 'page1ContinueButton').click()
        time.sleep(2) # wait for the page to load

        ## entering the details about the construction of the house
        # building type
        Select(driver.find_element(By.ID, "buildingType")).select_by_value(data["Building_type"])

        # construction type
        Select(driver.find_element(By.ID, "constructionType")).select_by_value(data["Construction_type"])

        # roof type
        Select(driver.find_element(By.ID, "roofType")).select_by_value(data["Roof_type"])
        
        # says whether or not there is a signnificant hazard risk
        if test_home_data_df.loc[person_i, "HazardRisk"].upper() == "YES":
            return "Webiste Does Not Quote For This House/ Person, Issue With HazardRisk Column"
        elif test_home_data_df.loc[person_i, "HazardRisk"].upper() == "NO":
            driver.find_element(By.XPATH, '//*[@id="homeDetails"]/div[2]/div[4]/fieldset/div[2]/ul/li[2]/label').click()
        else:
            print("Invalid Input Data Error: HazardRisk Column", end=" -- ")
            return "Invalid Input Data Error: HazardRisk Column"


        # standard of design and build
        driver.find_element(By.ID, test_home_data_df.loc[person_i, "HouseStandard"]).click()

        # shape of the land (how steep of a slope)
        driver.find_element(By.ID, test_home_data_df.loc[person_i, "LandShape"]).click()

        # year the house was built
        error_message = send_keys_data_entry("yearBuilt", "YearBuilt")
        if error_message != None:
            return error_message

        # year the person purchase the house
        error_message = send_keys_data_entry("yearPurchase", "PurchaseYear")
        if error_message != None:
            return error_message

        # the number of stories in the house
        Select(driver.find_element(By.ID, "numberOfStoreys")).select_by_value(data["Num_stories"])

        # the floor area of the house
        error_message = send_keys_data_entry("dwellingFloorArea", "DwellingFloorArea")
        if error_message != None:
            return error_message

        # does it have a freestanding garage
        if test_home_data_df.loc[person_i, "HasGarage"].upper() == "YES":
            driver.find_element(By.XPATH, '//*[@id="hasGarageContainer"]/fieldset/div[2]/ul/li[1]/label').click()
            error_message = send_keys_data_entry("garageFloorArea", "GarageFloorArea")
            if error_message != None:
                return error_message
        else: # has NO garage
            driver.find_element(By.XPATH, '//*[@id="hasGarageContainer"]/fieldset/div[2]/ul/li[2]/label').click()

        # Do you need cover for 'additional features' that are above a threshold value
        if test_home_data_df.loc[person_i, "AdditionalFeaturesAboveLimit"].upper() == "NO":
            Wait3.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="specialFeaturesContainer"]/fieldset/div[2]/ul/li[2]/label'))).click()
        else:
            print("Webiste Does Not Quote For This House/ Person, Issue With AdditionalFeaturesAboveLimit Column", end=" -- ")
            return "Webiste Does Not Quote For This House/ Person, Issue With AdditionalFeaturesAboveLimit Column" # if the house has additional features above the allowed limits

        # entering the Sum insured
        error_message = send_keys_data_entry("bldgSumInsured", "SumInsured")
        if error_message != None:
            return error_message

        # if you want to have to have glass excess
        if test_home_data_df.loc[person_i, "GlassExcess"].upper() == "YES":
            driver.find_element(By.XPATH, '//*[@id="glassBuyoutContainer"]/fieldset/div[2]/ul/li[1]/label').click()
        elif test_home_data_df.loc[person_i, "GlassExcess"].upper() == "NO":
            driver.find_element(By.XPATH, '//*[@id="glassBuyoutContainer"]/fieldset/div[2]/ul/li[2]/label').click()
        else:
            print("Invalid Input Data Error: GlassExcess Column", end=" -- ")
            return "Invalid Input Data Error: GlassExcess Column"
        

        ## clicking the button to 'Get your quote'
        driver.find_element(By.ID, "send-quote-id").click()
        time.sleep(5) # wait for the page to load


        ## selecting the excess on the policy
        # scraping all the excess options from the webpage
        excess_options = Wait10.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, '[for^="excess-option-item"]')))
        excess_values =[funct_defs.convert_money_str_to_int(option.get_attribute("for")) for option in excess_options]


        # choosing the excess value that is closest to our desired one
        chosen_excess = funct_defs.choose_excess_value(excess_values, int(test_home_data_df.loc[person_i, "Excess"]))

        # selecting the chosen excess value
        element = Wait10.until(EC.element_to_be_clickable((By.XPATH, f"//div[@id='excess-option{chosen_excess}']/span")))
        if element.get_attribute("class") != "checkmark checked": # if the box has not already been seleted, then click it
            element.click()

        # if the excess was changed, then handle it 
        if chosen_excess != int(test_home_data_df.loc[person_i, "Excess"]):
            # print and output an error message
            print(f"Excess Changed Warning! {test_home_data_df.loc[person_i, "Excess"]} not an option so {chosen_excess} selected", end=" -- ")
            output_df.loc[person_i, "AMI_Error_code"] = f"Excess Changed Warning! {test_home_data_df.loc[person_i, "Excess"]} not an option so {chosen_excess} selected"

            # set the excess for this example to be equal to the newly chosen excess
            output_df.loc[person_i, "Excess"] = chosen_excess


        time.sleep(3) # wait for page to load


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
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
    global driver, Wait3, Wait10
    driver, Wait3, Wait10 = funct_defs.load_webdriver()

"""
-------------------------
"""

### Defining the key functions for scraping from AA

# defining a function to scrape from the given company website (for an individual person/house)
def aa_home_premium_scrape(person_i):
    # defining a function which take the information from the spreadsheet and formats it so it can be used to scrape premium from aa website
    def aa_auto_data_format(person_i):

        # getting the persons birthdate out as a date object (allows us to get the correct format more easily)
        birthdate = test_home_data_df.loc[person_i,'DOB']

        # define a dict to store information for a given house/person
        data  = {'Has_permanent_residents': not test_home_data_df.loc[person_i,'Occupancy'] == 'Rented',
                'Street_address':f"{test_home_data_df.loc[person_i,'Street_number']} {test_home_data_df.loc[person_i,'Street_name']} {test_home_data_df.loc[person_i,'Street_type']}",
                "Birthdate_day":birthdate.strftime("%d"),
                "Birthdate_month":birthdate.strftime("%B"),
                "Birthdate_year":birthdate.strftime("%Y")
                 }
        return data


    # scrapes the insurance premium for a single vehicle/person at aa
    def aa_auto_scrape_premium(data):
        
        ## opening a window to get home insurance premium from
        if test_home_data_df.loc[person_i,'Occupancy'] == 'Unoccupied': # if the house is unoccupied
            # if the house is unoccupied, we must call in to get cover (so cannot be scraped)
            return "Doesn't Cover" 
        
        elif test_home_data_df.loc[person_i,'Occupancy'] in ['Owner occupied', 'Holiday home', 'Let to family/employees']: # if can be covered by standard home insurance
            # selects 'Home Insurance'
            Wait10.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="page"]/main/div[1]/div[2]/div[2]/div/div[1]/div[1]/div/ul/li[1]/a/div[2]/span'))).click() 

        elif test_home_data_df.loc[person_i,'Occupancy'] == 'Rented':
            # selects 'Landlord Insurance'
            Wait10.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="page"]/main/div[1]/div[2]/div[2]/div/div[1]/div[1]/div/ul/li[6]/a/div[2]/span'))).click() 


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
        driver.find_element(By.XPATH, '//*[@id="coverTypeButtons"]/label[1]/span').click()


        ## entering the house address
        # entering suburb + poscode + city
        suburb_input_box = driver.find_element(By.ID, "address.suburbPostcodeRegionCity")
        suburb_input_box.send_keys(str(test_home_data_df.loc[person_i,'Postcode'])) # entering the postcode into the input box
        Wait3.until(EC.element_to_be_clickable((By.XPATH, f'//*[@id="insuranceOptions"]/fieldset[4]/div[1]/div/ul/li[contains(text(), "{test_home_data_df.loc[person_i,'Suburb']}")]'))).click() # select the dropdown with the correct suburb


        # entering the street address (street name, type and number)
        street_address_input_box = driver.find_element(By.ID, "address.streetAddress")
        street_address_input_box.send_keys(data["Street_address"]) # entering the postcode into the input box



        # entering the policy holders data of birth
        Select(driver.find_element(By.ID, 'dateOfBirth-day')).select_by_value(data["Birthdate_day"])
        Select(driver.find_element(By.ID, 'dateOfBirth-month')).select_by_visible_text(data["Birthdate_month"])
        Select(driver.find_element(By.ID, 'dateOfBirth-year')).select_by_visible_text(data["Birthdate_year"])

        

        # waiting for testing purposes
        time.sleep(100000)

        # reformatting the montly and yearly premiums into integers
        monthly_premium, yearly_premium = funct_defs.convert_money_str_to_int("$0", cents=True), funct_defs. convert_money_str_to_int("$0", cents=True)

        # close the window we have just been scraping from
        driver.close()
        driver.switch_to.window(driver.window_handles[0]) # Switch back to the first tab (as the current tab we are on no longer exists)

        # returning the monthly/yearly premium and the adjusted agreed value
        return monthly_premium, yearly_premium


    # get time of start of each iteration
    start_time = time.time() 

    # run on the ith car/person
    try:
        # scrapes the insurance premium for a single house/person
        home_premiums = aa_auto_scrape_premium(aa_auto_data_format(person_i))

        if home_premiums != None and not isinstance(home_premiums, str): # if home_premiums is the scraped premiums (NOT just an error message)

            # print the scraping results
            print(home_premiums[0], home_premiums[1], end =" -- ")

            # save the scraped premiums to the output dataset
            output_df.loc[person_i, "Tower_monthly_premium"] = home_premiums[0] # monthly
            output_df.loc[person_i, "Tower_yearly_premium"] = home_premiums[1] # yearly

        # all these are processing error codes
        elif home_premiums == "Doesn't Cover":
            output_df.loc[person_i, "Tower_Error_code"] = "Website Does Not Quote For This Car House/Person"
        elif home_premiums == "Unable to find car variant":
            output_df.loc[person_i, "Tower_Error_code"] = "Unable to find car variant"
        elif "Invalid Input Data Error" in home_premiums:
            output_df.loc[person_i, "Tower_Error_code"] = home_premiums
        else:
            output_df.loc[person_i, "Tower_Error_code"] = "Unknown Error"

    except:
        try: # checks if the reason our code failed is because the 'we need more information' pop up appeareds
            Wait3.until(EC.visibility_of_element_located( (By.XPATH, "//*[@id='ui-id-3' and text() = 'We need more information']") ) )
            print("Need more information", end= " -- ")
            output_df.loc[person_i, "Tower_Error_code"] = "Webiste Does Not Quote For This Car Variant/ Person"
        except exceptions.TimeoutException:
            raise Exception("View Errors")
            print("Unknown Error!!", end= " -- ")
            output_df.loc[person_i, "Tower_Error_code"] = "Unknown Error"


    end_time = time.time() # get time of end of each iteration
    print("Elapsed time:", round(end_time - start_time, 2)) # print out the length of time taken
    

### a function to scrape premiums all given examples on aa's website
def aa_auto_scape_all():
    # define which row incidies to scrape from
    input_indexes = funct_defs.read_indicies_to_scrape()

    # Open the webpage AA webpage
    driver.get("https://www.aainsurance.co.nz/")
    Wait3.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="page"]/main/div[1]/div[2]/div[2]/div/div[1]/div[1]/div/div/div/div[1]'))).click() # click the 'Get a quote button'

    # loop through all examples in test_home_data spreadsheet
    for person_i in input_indexes: 

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


    #funct_defs.export_auto_dataset(input_indexes)




def main():
    # performing all data reading in and preprocessing
    global test_home_data_df, output_df
    test_home_data_df, output_df = funct_defs.dataset_preprocess("AA")

    # loading the webdriver
    load_webdriver()

    # scrape all of the insurance premiums for the given cars from aa
    aa_auto_scape_all()

    # Close the browser window
    #driver.quit()

main()
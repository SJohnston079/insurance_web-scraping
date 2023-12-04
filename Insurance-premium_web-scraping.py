# webscraping related imports
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

# data management relates imports
import pandas as pd
from datetime import datetime

# reads in the test data for car insurance inputs
test_auto_data = pd.read_excel("S:\\Library\\IQS\\Test data\\test_auto_data.xlsx")

# defining a function which take the information from the spreadsheet and formats it so it can be used to scrape premium from ami website
def ami_auto_data_format(person_i):
    # getting the address in the correct format
    address = ""

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
    elif drivers_license_years >= 5:
        drivers_license_years = "5 years or more"
    else: # for for generic 'International' (non-NZ) licence
        drivers_license_years = "{} years".format(drivers_license_years)

    # define a dict to store information for a given person and car for ami
    ami_data = {"Manufacturer":test_auto_data.loc[person_i,'Manufacturer'],
                "Model":test_auto_data.loc[person_i,'Model'],
                "Vehicle_year":test_auto_data.loc[person_i,'Vehicle_year'],
                "Body_type":test_auto_data.loc[person_i,'Body'].upper(),
                "Engine_size":"{}cc/{}L".format(int(test_auto_data.loc[0,'CC']), round(test_auto_data.loc[0,'CC']/1000, 1)),
                "Business_use":test_auto_data.loc[person_i,'BusinessUser'],
                "Full_garaging_address":"TODO",
                "birthdate_day":int(birthdate.strftime("%d")),
                "birthdate_month":birthdate.strftime("%B"),
                "birthdate_year":int(birthdate.strftime("%Y")),
                "sex":test_auto_data.loc[person_i,'Gender'],
                "drivers_license_type":drivers_license_type,
                "drivers_license_years":drivers_license_years, # years since driver got their learners licence
                "incidents_5_year":test_auto_data.loc[person_i,'Incidents_last5years_AMISTATE'],
                }
    
    # adding info on the date and type of incident to the ami_data dictionary ONLY if the person has had an incident within the last 5 years
    incident_date = test_auto_data.loc[person_i,'Date_of_incident']
    if ami_data["incidents_5_year"] == "Yes":
        ami_data["incident_date_month"] = incident_date.strftime("%B")
        ami_data["incident_date_year"] = int(incident_date.strftime("%Y"))
        incident_type = test_auto_data.loc[person_i,'Type_incident'].split("-")[0].strip()
        if incident_type == "Not at fault":
            ami_data["incident_type"] = "Not At Fault Accident"
        else:
            ami_data["incident_type"] = "At Fault Accident"

    # returns the dict object containing all the formatted data
    return ami_data

print(ami_auto_data_format(0))

# loads chromedriver
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))

# define the implicit wait time for the session
driver.implicitly_wait(10)

# scrapes the insurance premium for a single vehicle and person at ami
def ami_scrape_premium(data):
    # Open the webpage
    driver.get("https://secure.ami.co.nz/css/car/step1")

    # Find the button "Make, Model, Year" and click it
    make_model_year_button = driver.find_element(By.ID, "ie_MMYPrepareButton")
    make_model_year_button.click()


    # inputting the car manufacturer

    car_manfacturer_element = driver.find_element(By.ID, "vehicleManufacturer") # find car manufacturer input box
    car_manfacturer_element.click() # open the input box
    time.sleep(2)
    car_manfacturer_element.send_keys(data["Manufacturer"]) # input the company that manufactures the car
    time.sleep(2)
    driver.find_element(By.XPATH, "//a[@class='ui-corner-all' and text()='{}']".format(data["Manufacturer"])).click() # clicking the button which has the correct manufacturer information
    if car_manfacturer_element.get_attribute('oldvalue') != data["Manufacturer"]: # checking that correct manufactuerer selected
        return False
    time.sleep(1) # wait for page to process information


    # inputting car model
    driver.find_element(By.ID, "Model").click() # find car model input box and open it
    driver.find_element(By.XPATH, "//div[text()='{}']".format(data["Model"])).click() # clicking the button which has the correct car model information
    time.sleep(1) # wait for page to process information

    # inputting car year
    driver.find_element(By.ID, "Year").click() # find car year input box and open it
    driver.find_element(By.XPATH, "//div[text()='{}']".format(data["Vehicle_year"])).click() # clicking the button which has the correct car model information
    time.sleep(1) # wait for page to process information

    # inputting car body type
    driver.find_element(By.ID, "BodyType").click() # find car BodyType input box and open it
    driver.find_element(By.XPATH, "//div[text()='{}']".format(data["Body_type"])).click() # clicking the button which has the correct car model information
    time.sleep(1) # wait for page to process information

    # inputting car engine size
    
    driver.find_element(By.ID, "EngineSize").click() # find car BodyType input box and open it
    driver.find_element(By.XPATH, "//div[text()='{}']".format(data["Engine_size"])).click() # clicking the button which has the correct car model information
    time.sleep(1) # wait for page to process information
    
    # select the 1st vehicle in the final options 
    driver.find_element(By.ID, "searchedVehicleSpan_0").click() # click button to select final vehicle option



    # selects whether or not the car is used for business
    driver.find_element(By.ID, "bIsBusinessUse_false").click() # clicks "False" button


    # inputs the address the car is kept at
    driver.find_element(By.ID, "garagingAddress_fullAddress").send_keys("79 PROSPECT TERRACE, MOUNT EDEN, AUCKLAND 1024") # inputs the address to the given field
    driver.find_element(By.XPATH, "//a[text()='79 PROSPECT TERRACE, MOUNT EDEN, AUCKLAND 1024']").click()


    # inputs driver birth date
    driver.find_element(By.ID, "driverDay_1").send_keys("12") # input day
    driver.find_element(By.ID, "driverMonth_1").send_keys("February") # input month
    driver.find_element(By.ID, "driverYear_1").send_keys("2003") # input year

    # select driver sex
    driver.find_element(By.ID, "male_1").click() # selects male



    # enter drivers licsence info
    driver.find_element(By.ID, "DriverLicenceType_1").click() # open the drivers license type options box
    driver.find_element(By.XPATH, "//div[text()='NZ Full']").click() # select the drivers license type
    driver.find_element(By.ID, "DriverYearsOfDriving_1").click() # open years since got learners box
    driver.find_element(By.XPATH, "//div[text()='3 years']").click() # select correct years since got learners

    # input if there have been any indicents
    driver.find_element(By.NAME, "driverLoss").click() # clicks button saying that you have had an incident
    driver.find_element(By.ID, "DriverIncidentType_1").click() # opens incident type option box
    driver.find_element(By.XPATH, "//div[text()='Not At Fault Accident']").click() # selects the driver incident type

    driver.find_element(By.ID, "DriverIncidentMonth_1").click() # opens incident type option box
    driver.find_element(By.XPATH, "//html//body//div[15]//div//div[text()='March']").click() # selects the driver incident type
    driver.find_element(By.ID, "DriverIncidentYear_1").click() # opens incident type option box
    driver.find_element(By.XPATH, "//html//body//div[16]//div[text()='2021']").click() # selects the driver incident type


    time.sleep(2) # wait a bit

    # click button to get quote
    driver.find_element(By.ID, "quoteSaveButton").click()

    time.sleep(5)

    # delete all cookies to reset the page
    driver.delete_all_cookies()

    # if scrape is successful return true
    return True


ami_scrape_premium(ami_auto_data_format(0))

# Close the browser window
#driver.quit()

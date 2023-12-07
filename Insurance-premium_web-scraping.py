# webscraping related imports
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# data management relates imports
import pandas as pd
from datetime import datetime

# defining a function that will scrape all of the ami cars
def ami_auto_scrape_all():
    # defining a function which take the information from the spreadsheet and formats it so it can be used to scrape premium from ami website
    def ami_auto_data_format(person_i):
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
                    "Model_type":test_auto_data.loc[person_i,'Type'],
                    "Vehicle_year":test_auto_data.loc[person_i,'Vehicle_year'],
                    "Body_type":test_auto_data.loc[person_i,'Body'].upper(),
                    "Engine_size":"{}cc/{}L".format(int(test_auto_data.loc[person_i,'CC']), round(test_auto_data.loc[person_i,'CC']/1000, 1)),
                    "Immobiliser":test_auto_data.loc[person_i,'Immobiliser_alarm'],
                    "Business_use":test_auto_data.loc[person_i,'BusinessUser'],
                    "Unit":test_auto_data.loc[person_i,'Unit_number'],
                    "Street_number":test_auto_data.loc[person_i,'Street_number'],
                    "Street_name":test_auto_data.loc[person_i,'Street_name'] + " " + test_auto_data.loc[person_i,'Street_type'],
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

        # Find the button "Make, Model, Year" and click it
        make_model_year_button = driver.find_element(By.ID, "ie_MMYPrepareButton")
        make_model_year_button.click()


        # inputting the car manufacturer
        car_manfacturer_element = driver.find_element(By.ID, "vehicleManufacturer") # find car manufacturer input box
        car_manfacturer_element.click() # open the input box
        time.sleep(1) # wait for page to process information
        car_manfacturer_element.send_keys(data["Manufacturer"]) # input the company that manufactures the car
        time.sleep(1) # wait for page to process information
        driver.find_element(By.XPATH, "//a[@class='ui-corner-all' and text()='{}']".format(data["Manufacturer"])).click() # clicking the button which has the correct manufacturer information
        if car_manfacturer_element.get_attribute('oldvalue') != data["Manufacturer"]: # checking that correct manufactuerer selected
            return False

        # inputting car model
        if driver.find_element(By.XPATH, "/html/body/div[4]/form/div/div[1]/div/div[1]/div/fieldset/ul[1]/li[2]/div/span/input[2]").get_attribute("value") != data["Model"]:
            Wait.until(EC.element_to_be_clickable((By.ID, "Model"))).click() # wait until car model input box is clickable, then open it
            Wait.until(EC.element_to_be_clickable((By.XPATH, "//div[text()='{}']".format(data["Model"])))).click() # wait until button which has the correct car model information is clickable, then click

        # inputting car year
        if driver.find_element(By.XPATH, "/html/body/div[4]/form/div/div[1]/div/div[1]/div/fieldset/ul[1]/li[3]/div/span/input[2]").get_attribute("value") != data["Vehicle_year"]:
            Wait.until_not(lambda x: x.find_element(By.ID, "searchByMMYLoading").is_displayed()) # wait until the "loading element" is not being displayed
            time.sleep(0.5)
            driver.find_element(By.ID, "Year").click() # find car year input box then open it
            Wait.until(EC.element_to_be_clickable((By.XPATH, "//div[text()='{}']".format(data["Vehicle_year"])))).click() # clicking the button which has the correct car model information

        # inputting car body type
        if driver.find_element(By.XPATH, "/html/body/div[4]/form/div/div[1]/div/div[1]/div/fieldset/ul[1]/li[4]/div/span/input[2]").get_attribute("value") != data["Body_type"].upper():
            Wait.until_not(lambda x: x.find_element(By.ID, "searchByMMYLoading").is_displayed()) # wait until the "loading element" is not being displayed
            driver.find_element(By.ID, "BodyType").click() # find car BodyType input box and open it
            driver.find_element(By.XPATH, "//div[text()='{}']".format(data["Body_type"])).click() # clicking the button which has the correct car model information

        # inputting car engine size
        if driver.find_element(By.XPATH, "/html/body/div[4]/form/div/div[1]/div/div[1]/div/fieldset/ul[1]/li[5]/div/span/input[2]").get_attribute("value") != data["Engine_size"]:
            Wait.until_not(lambda x: x.find_element(By.ID, "searchByMMYLoading").is_displayed()) # wait until the "loading element" is not being displayed
            driver.find_element(By.ID, "EngineSize").click() # find car BodyType input box and open it
            driver.find_element(By.XPATH, "//div[text()='{}']".format(data["Engine_size"])).click() # clicking the button which has the correct car model information
        time.sleep(1) # wait for page to process information
        
        # select the final vehicle option
        if not pd.isna(data["Model_type"]):
            driver.find_element(By.XPATH, "//label[contains(text(), '{}')]".format(data["Model_type"])).click() # click button to select final vehicle option
        else:
            driver.find_element(By.ID, "searchedVehicleSpan_0").click() # click button to select final vehicle option

        # selects whether or not the car has an immobiliser
        try:
            if data["Immobiliser"] == "Yes":
                driver.find_element(By.ID, "bHasImmobilizer_true").click() # clicks True button
            else:
                driver.find_element(By.ID, "bHasImmobilizer_false").click() # clicks False button
        except:
            pass

        # selects whether or not the car is used for business
        if data["Business_use"] == "No":
            driver.find_element(By.ID, "bIsBusinessUse_false").click() # clicks "False" button
        else:
            driver.find_element(By.ID, "bIsBusinessUse_true").click() # clicks "False" button


        # inputs the address the car is kept at
        driver.find_element(By.ID, "garagingAddress_autoManualRadio").click() # click button to enter address manually
        if not pd.isna(data["Unit"]): driver.find_element(By.ID, "garagingAddress_manualUnitNumber").send_keys(data["Unit"]) # input Unit/Apt IF is applicable
        driver.find_element(By.ID, "garagingAddress_manualStreetNumber").send_keys(data["Street_number"])
        driver.find_element(By.ID, "garagingAddress_manualStreetName").send_keys(data["Street_name"])
        try:
            time.sleep(2) # wait for options list to pop up
            driver.find_element(By.XPATH, "//*[contains(text(), '{}') or contains(text(),'{}')]".format(data["Suburb"], data["Postcode"])).click()
        except:
            try:
                driver.find_element(By.XPATH, "//li[@class='ui-menu-item']//a[contains(text(), '{}')]".format(data["Postcode"])).click()
            except:
                driver.find_element(By.ID, "garagingAddress_manualSuburb").send_keys(data["Suburb"])
                driver.find_element(By.XPATH, "//li[@class='ui-menu-item']//a[contains(text(), '{}')]".format(data["Postcode"])).click()
        
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

        # checks if 'we need more information' pop up appears

        try:
            # wait until next page is loaded
            annual_risk_premium = Wait.until(EC.presence_of_element_located((By.ID, "annualRiskPremium")))

            # scrape the premium
            monthy_premium = float(driver.find_element(By.ID, "dollars").text.replace(",", "") + driver.find_element(By.ID, "cents").text)
            yearly_premium = float(annual_risk_premium.text.replace(",", "")[1:])

            # return the scraped premiums
            return monthy_premium, yearly_premium
        except:
            return -1, -1 # return monthy and yearly premiums as -1 to indicate the premiums cannot be scraped


    # loop through all cars in test spreadsheet
    for person_i in range(43, len(test_auto_data)):
        start_time = time.time() # get time of start of each iteration

        # run on the ith car/person
        monthy_premium, yearly_premium = ami_auto_scrape_premium(ami_auto_data_format(person_i))
        print(person_i, ":", monthy_premium, yearly_premium, end =" -- ")
        """
        try:
            # run on the ith car/person
            monthy_premium, yearly_premium = ami_auto_scrape_premium(ami_auto_data_format(person_i))
            print(person_i, ":", monthy_premium, yearly_premium, end =" -- ")
        except:
            # print out the formatted data (for debugging)
            print(person_i, ":", ami_auto_data_format(person_i), end=" -- ")
        """
        end_time = time.time() # get time of end of each iteration
        print("Elapsed time:", round(end_time - start_time,2)) # print out the length of time taken

def main():
    # reads in the test data for car insurance inputs
    global test_auto_data
    test_auto_data = pd.read_excel("C:\\Users\\samuel.johnston\\Documents\\Insurance_web-scraping\\test_auto_data1.xlsx", dtype={"Postcode":"int"})
    #pd.to_numeric(test_auto_data.loc[:,"Postcode"].astype(int)) (when I add all postcode values to the excel spreadsheet add back in to convert all postcodes from floats to ints)

    # loads chromedriver
    global driver # defines driver as a global variable
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))

    # define the implicit wait time for the session
    driver.implicitly_wait(2)

    # defines Wait, a dynamic wait template
    global Wait
    Wait = WebDriverWait(driver, 10)

    # scrape all of the insurance premiums for the given cars on ami
    ami_auto_scrape_all()

    # Close the browser window
    driver.quit()
main()
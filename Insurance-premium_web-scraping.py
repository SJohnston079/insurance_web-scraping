import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

# loads chromedriver
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))

# define the implicit wait time for the session
driver.implicitly_wait(10)


# Open the webpage
driver.get("https://secure.ami.co.nz/css/car/step1")

# Find the button "Make, Model, Year" and click it
make_model_year_button = driver.find_element(By.ID, "ie_MMYPrepareButton")
make_model_year_button.click()


# inputting the car manufacturer
car_manfacturer_element = driver.find_element(By.ID, "vehicleManufacturer") # find car manufacturer input box
car_manfacturer_element.click() # open the input box
time.sleep(2)
car_manfacturer_element.send_keys("Toyota") # input the info
time.sleep(2)
driver.find_element(By.XPATH, "//a[@class='ui-corner-all' and text()='Toyota']").click() # clicking the button which has the correct manufacturer information
if car_manfacturer_element.get_attribute('oldvalue') != "Toyota": # checking that correct manufactuerer selected
    print("Error, incorrect manufacturer selected")
    driver.close()
time.sleep(2) # wait for page to process information


# inputting car model
driver.find_element(By.ID, "Model").click() # find car model input box and open it
driver.find_element(By.XPATH, "//div[text()='Camry']").click() # clicking the button which has the correct car model information
time.sleep(2) # wait for page to process information

# inputting car year
driver.find_element(By.ID, "Year").click() # find car year input box and open it
driver.find_element(By.XPATH, "//div[text()='2022']").click() # clicking the button which has the correct car model information
time.sleep(2) # wait for page to process information

# inputting car body type
driver.find_element(By.ID, "BodyType").click() # find car BodyType input box and open it
driver.find_element(By.XPATH, "//div[text()='SEDAN']").click() # clicking the button which has the correct car model information
time.sleep(2) # wait for page to process information

# inputting car engine size
driver.find_element(By.ID, "EngineSize").click() # find car BodyType input box and open it
driver.find_element(By.XPATH, "//div[text()='2487cc/2.5L']").click() # clicking the button which has the correct car model information
time.sleep(2) # wait for page to process information

# select the 1st vehicle in the final options 
driver.find_element(By.ID, "searchedVehicleSpan_0").click() # click button to select final vehicle option



# selects whether or not the car is used for business
driver.find_element(By.ID, "bIsBusinessUse_false").click() # clicks "False" button


# inputs the address the car is kept at
driver.find_element(By.ID, "garagingAddress_fullAddress").send_keys("79 PROSPECT TERRACE, MOUNT EDEN, AUCKLAND 1024") # inputs the address to the given field
driver.find_element(By.XPATH, "//a[text()='79 PROSPECT TERRACE, MOUNT EDEN, AUCKLAND 1024']").click()


# inputs driver age
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
driver.find_element(By.XPATH, "//div[text()='March']").click() # selects the driver incident type
driver.find_element(By.ID, "DriverIncidentYear_1").click() # opens incident type option box
driver.find_element(By.XPATH, "//html//body//div[16]//div[text()='2021']").click() # selects the driver incident type


time.sleep(2) # wait a bit

# click button to get quote
driver.find_element(By.ID, "quoteSaveButton").click()


time.sleep(1000)

# Close the browser window
driver.quit()
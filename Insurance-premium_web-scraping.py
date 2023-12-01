import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))




# Open the webpage
driver.get("https://secure.ami.co.nz/css/car/step1")

# Find the button "Make, Model, Year" and click it
make_model_year_button = driver.find_element(By.ID, "ie_MMYPrepareButton")
make_model_year_button.click()

# finding the elements to input car info

# inputting the car manufacuter
car_make_element = driver.find_element(By.ID, "vehicleManufacturer")
car_make_element.send_keys("Toyota")
chosen_car_manufacturer_element = driver.find_element(By.XPATH, "/html/body/div[4]/form/div/div[1]/div/div[1]/div/fieldset/ul[2]/li/a")
chosen_car_manufacturer_element.click()


# inputting the car make (e.g. Camry for a Toyota make)
#element = driver.find_element(By.ID, "Model").click()
#car_make_element = driver.find_element(By.XPATH, "/html/body/div[5]/div/div[25]")
#car_make_element.click()

time.sleep(1000)

# Close the browser window
driver.quit()
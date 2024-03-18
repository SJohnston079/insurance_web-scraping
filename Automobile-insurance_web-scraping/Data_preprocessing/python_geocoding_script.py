"""
Python script for batch geocoding of addresses using the Google Geocoding API.
This script allows for massive lists of addresses to be geocoded for free by pausing when the 
geocoder hits the free rate limit set by Google (2500 per day).  If you have an API key for paid
geocoding from Google, set it in the API key section.
Addresses for geocoding can be specified in a list of strings "addresses". In this script, addresses
come from a csv file with a column "Address". Adjust the code to your own requirements as needed.
After every 500 successul geocode operations, a temporary file with results is recorded in case of 
script failure / loss of connection later.
Addresses and data are held in memory, so this script may need to be adjusted to process files line
by line if you are processing millions of entries.
Shane Lynn
5th November 2016
"""

import pandas as pd
import requests
import logging
import time

logger = logging.getLogger("root")
logger.setLevel(logging.DEBUG)
# create console handler
ch = logging.StreamHandler()
ch.setLevel(logging.DEBUG)
logger.addHandler(ch)

#------------------ CONFIGURATION -------------------------------

# Set your Google API key here. 
# Even if using the free 2500 queries a day, its worth getting an API key since the rate limit is 50 / second.
# With API_KEY = None, you will run into a 2 second delay every 10 requests or so.
# With a "Google Maps Geocoding API" key from https://console.developers.google.com/apis/, 
# the daily limit will be 2500, but at a much faster rate.
# Example: API_KEY = 'AIzaSyC9azed9tLdjpZNjg2_kVePWvMIBq154eA'
API_KEY = 'AIzaSyCkWdxhcs4A6gFloNCFUqmg1WFj1lpplOQ'
# Region ccTLD, set to None if for worldwide address, otherwise a ccTLD for a region ('nz' for New Zealand)
region_ccTLD = 'nz'
# Backoff time sets how many minutes to wait between google pings when your API limit is hit
BACKOFF_TIME = 1/60
# Set your output file name here.
output_filename = 'C:\\Users\\samuel.johnston\\Documents\\Insurance_web-scraping\\test_auto_data-postcodes.csv'
# Set your input file here
input_filename = "C:\\Users\\samuel.johnston\\Documents\\Insurance_web-scraping\\test_auto_data1.xlsx"
# Specify the column name in your input data that contains addresses here
address_column_name = "Full Address"
# Return Full Google Results? If True, full JSON results from Google are included in output
RETURN_FULL_RESULTS = True

#------------------ DATA LOADING --------------------------------

# Read the data to a Pandas Dataframe
#data = pd.read_csv(input_filename, encoding='latin-1') #data = pd.read_csv(input_filename, encoding='utf8') # if error here use # data = pd.read_csv(input_filename, encoding='latin-1')
data = pd.read_excel(input_filename, sheet_name="Postcode mapping")

if address_column_name not in data.columns: 
	raise ValueError("Missing Address column in input data")

# Form a list of addresses for geocoding:
# Make a big list of all of the addresses to be processed.
addresses = data[address_column_name].tolist()

#------------------	FUNCTION DEFINITIONS ------------------------

def get_google_results(address, region=None, api_key=None, return_full_response=False):
    """
    Get geocode results from Google Maps Geocoding API.
    
    Note, that in the case of multiple google geocode reuslts, this function returns details of the FIRST result.
    
    @param address: String address as accurate as possible. For Example "18 Grafton Street, Dublin, Ireland"
    @param api_key: String API key if present from google. 
                    If supplied, requests will use your allowance from the Google API. If not, you
                    will be limited to the free usage of 2500 requests per day.
    @param return_full_response: Boolean to indicate if you'd like to return the full response from google. This
                    is useful if you'd like additional location details for storage or parsing later.
    """
    # Set up your Geocoding url
    geocode_url = "https://maps.googleapis.com/maps/api/geocode/json?address={}".format(address)
    if region_ccTLD is not None:
      geocode_url = geocode_url + "&region={}".format(region_ccTLD)
	
    if api_key is not None:
      geocode_url = geocode_url + "&key={}".format(api_key)
        
    # Ping google for the reuslts:
    results = requests.get(geocode_url)
    # Results will be in JSON format - convert to dict using requests functionality
    results = results.json()
    
    # if there's no results or an error, return empty results.
    if len(results['results']) == 0:
        output = {
            "street_number" : None,
            "street" : None,
            "suburb" : None,
            "town_city" : None,
            "region" : None,
            "country" : None,
            "postal_code" : None,
            "formatted_address" : None,
            "latitude": None,
            "longitude": None,
            "accuracy": None,
            "google_place_id": None,
            "type": None
        }
    else:    
        answer = results['results'][0]
        output = {
            "street_number" : ",".join([x['long_name'] for x in answer.get('address_components') 
                                        if 'street_number' in x.get('types')]),
            "street" : ",".join([x['long_name'] for x in answer.get('address_components') 
                                        if 'route' in x.get('types')]),
            "suburb" : ",".join([x['long_name'] for x in answer.get('address_components') 
                                        if 'sublocality_level_1' in x.get('types')]),
            "town_city" : ",".join([x['long_name'] for x in answer.get('address_components') 
                                        if 'locality' in x.get('types')]),
            "region" : ",".join([x['long_name'] for x in answer.get('address_components') 
                                        if 'administrative_area_level_1' in x.get('types')]),
            "country" : ",".join([x['long_name'] for x in answer.get('address_components') 
                                        if 'country' in x.get('types')]),
            "postal_code": ",".join([x['long_name'] for x in answer.get('address_components') 
                                        if 'postal_code' in x.get('types')]),
            "formatted_address" : answer.get('formatted_address'),
            "latitude": answer.get('geometry').get('location').get('lat'),
            "longitude": answer.get('geometry').get('location').get('lng'),
            "accuracy": answer.get('geometry').get('location_type'),
            "google_place_id": answer.get("place_id"),
            "type": ",".join(answer.get('types'))
            
        }
        
    # Append some other details:    
    output['input_string'] = address
    output['number_of_results'] = len(results['results'])
    output['status'] = results.get('status')
    if return_full_response is True:
        output['response'] = results
    
    return output

#------------------ PROCESSING LOOP -----------------------------

# Ensure, before we start, that the API key is ok/valid, and internet access is ok
test_result = get_google_results("London, England", None, API_KEY, False)#RETURN_FULL_RESULTS)
if (test_result['status'] != 'OK') or (test_result['formatted_address'] != 'London, UK'):
    logger.warning("There was an error when testing the Google Geocoder.")
    raise ConnectionError('Problem with test results from Google Geocode - check your API key and internet connection.')

# Create a list to hold results
results = []
# Go through each address in turn
for address in addresses:
    # While the address geocoding is not finished:
    geocoded = False
    while geocoded is not True:
        # Geocode the address with google
        try:
            geocode_result = get_google_results(address, region_ccTLD, API_KEY, return_full_response=RETURN_FULL_RESULTS)
        except Exception as e:
            logger.exception(e)
            logger.error("Major error with {}".format(address))
            logger.error("Skipping!")
            geocoded = True
            
        # If we're over the API limit, backoff for a while and try again later.
        if geocode_result['status'] == 'OVER_QUERY_LIMIT':
            logger.info("Hit Query Limit! Backing off for a bit.")
            time.sleep(BACKOFF_TIME * 60) # sleep for 30 minutes
            geocoded = False
        else:
            # If we're ok with API use, save the results
            # Note that the results might be empty / non-ok - log this
            if geocode_result['status'] != 'OK':
                logger.warning("Error geocoding {}: {}".format(address, geocode_result['status']))
            logger.debug("Geocoded: {}: {}".format(address, geocode_result['status']))
            results.append(geocode_result)           
            geocoded = True

    # Print status every 100 addresses
    if len(results) % 100 == 0:
    	logger.info("Completed {} of {} address".format(len(results), len(addresses)))
            
    # Every 500 addresses, save progress to file(in case of a failure so you have something!)
    if len(results) % 500 == 0:
        pd.DataFrame(results).to_csv("{}_bak".format(output_filename))

# All done
logger.info("Finished geocoding all addresses")
# Write the full results to csv using the pandas library.
pd.DataFrame(results).to_csv(output_filename, encoding='utf8')

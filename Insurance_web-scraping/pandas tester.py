import pandas as pd
from datetime import datetime, date
import os

'''
# Get the absolute path of the current Python file
file_path = os.path.abspath(__file__)

# Get the directory of the current Python file
file_directory = os.path.dirname(file_path)

# set the working directory to be the directory of this file
os.chdir(file_directory)


# performing the data reading in and preprocessing
def dataset_preprocess():
    # read in the data

    test_auto_data_df = pd.read_csv("test_auto_data1.csv", dtype={"Postcode":"int"})

    # sets all values of the policy start date to be today's date
    for key in test_auto_data_df:
        test_auto_data_df['PolicyStartDate'] = datetime.strftime(date.today(), "%d/%m/%Y")

    # creates a new dataframe to save the scraped info
    global aa_output_df
    aa_output_df = test_auto_data_df.loc[:, ["Sample Number", "PolicyStartDate"]]
    aa_output_df["AA_agreed_value"] = test_auto_data_df["AgreedValue"].to_string(index=False).strip().split()
    aa_output_df["AA_monthly_premium"] = ["-1"] * len(test_auto_data_df)
    aa_output_df["AA_yearly_premium"] = ["-1"] * len(test_auto_data_df)
    aa_output_df["AA_agreed_value_minimum"] = [-1] * len(test_auto_data_df)
    aa_output_df["AA_agreed_value_maximum"] = [-1] * len(test_auto_data_df)

dataset_preprocess()


insurance_premium_web_scraping_AA_df = pd.read_csv("Individual-company_data-files\\ami_scraped_auto_premiums.csv")
insurance_premium_web_scraping_AA_df.set_index("Sample Number", drop=True, inplace=True)
aa_output_df.set_index("Sample Number", drop=True, inplace=True)

auto_dataset_for_export = insurance_premium_web_scraping_AA_df.combine_first(aa_output_df.iloc[8:9])

print(insurance_premium_web_scraping_AA_df)
print(aa_output_df)
print(auto_dataset_for_export)
'''

#input_indexes = str([1,2,3,4])
#input_indexes = input_indexes.replace("[", "").replace("]", "").split(",")
#input_indexes = list(map(int, input_indexes))

print(str(range(0, 1000)))
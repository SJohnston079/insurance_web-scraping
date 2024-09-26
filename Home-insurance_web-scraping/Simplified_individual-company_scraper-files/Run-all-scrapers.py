# required library imports
import subprocess as sp
import multiprocessing
import funct_defs





# get the directory of this file
directory = funct_defs.get_current_file_directory()




# a function that calls all of the individual insurance company files
def run_subprocess(company):
    # running the process (the python file)
    sp.Popen(['python', f'{directory}\\Insurance_premium_web_scraping_{company}.py'], stdin=sp.PIPE, text=True)




def main():
    insurance_companies = ["AA", "AMI", "TOWER"]
    with multiprocessing.Pool() as pool:
        pool.map(run_subprocess, insurance_companies)




# run main() and ensure that it is only run when the code is called directly
if __name__ == "__main__":
    main()
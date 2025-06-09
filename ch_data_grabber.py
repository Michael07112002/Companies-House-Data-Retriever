# ToDo: 
# Either recreate spreedsheet so headers are headers or find way to so with pandas
# Find way to search linked in api 
# Scrape company accounts



# Checks 

# test main on all companies 
# Then output the df to csv or excel file
# Current estimated run time for program is 312.5 mins (5 hrs 12.5 mins)

# Libraries required
import requests 
from requests.auth import HTTPBasicAuth 
import json
import pandas as pd 
import numpy as np 
import re
from googlesearch import search
import time

# Take inputted csv file 
# df = pd.read_csv("/mnt/c/users/micha/downloads/ME_Pre-fix-Primary_Targets.csv")
df = pd.read_csv("/mnt/c/users/micha/downloads/ME_Pre-fix-Incorp_after_07.04.2023.csv")




company_numbers_df = df[df["company_number"].notna()]
company_numbers = company_numbers_df["company_number"]

api_key = "97ca52aa-b171-4f22-8344-04bb2a255b2a" 


# Continues getting feeback for each function from chat gpt 
# get_directors: done 
# get_charge_holders: done

def main(): 
    # for i in range(5):
    for i in range(len(company_numbers)): 
        print(f"Company: {company_numbers[i]}")
        print(f"Company {i + 1} out of {len(company_numbers)}")
        print(f"Progress {(i+1)*100/len(company_numbers)}%")
        start = time.time()
        # Get director name and DoB and input into data frame
        director_info = get_directors(company_numbers[i]) 
        df.at[i, "Director Name"] = director_info["name"]
        df.at[i, "Director DoB"] = director_info["DoB"]
        time.sleep(0.5) # Pauses program for 30 seconds in order to not go over the 600 request per 5 mins limit on the Companies House api

        # Get the charge holders if they exist
        charge_holders = get_charge_holder(company_numbers[i])
        if charge_holders["status"] == "outstanding":
            df.at[i, "Charge Holders"] = charge_holders["charge_holders"]
        else: 
            df.at[i, "Charge Holders"] = charge_holders["status"]
        time.sleep(0.5)

        # Find out if accounts are dormant
        dormant = get_dormant_accounts(company_numbers[i])  
        df.at[i, "Dormant Latest Accounts?"] = dormant 
        time.sleep(0.5)

        # Find out if accounts and confirmation statement are overdue
        overdue = get_overdue(company_numbers[i])
        df.at[i, "Accounts Overdue?"] = overdue["accounts"]
        df.at[i, "Confirmation Statement Overdue?"] = overdue["confirmation statement"]
        time.sleep(0.5)
        end = time.time() 
        print(f"Run time for 1 company: {end - start}")
    
    # Create excel spreedsheet from dataframe
    df.to_excel("/mnt/c/users/micha/downloads/output_Incorp_after_07.04.2023.xlsx", sheet_name="Incorp after 07.04.2023", index=False)
    print("Excel sheet successfully created")

    



def get_directors(company_number):

    # Request info on officers
    url = f"https://api.company-information.service.gov.uk/company/{company_number}/officers" 
    response = requests.get(url, auth=HTTPBasicAuth(api_key, "")) 

    if response.status_code == 200: 
        data = response.json() 

        # Go through each officer to find the first one listed as director 
        officer_info = data.get("items", [])
        directors = {}
        for officer in officer_info: 
            # Using try-except to handle missing data (e.g. no date of birth or role listed) 
            try:
                
                

                # Check that they are a director
                if officer.get("officer_role", None) == "director": 
                    # Extract director's name and date of birth
                    name = officer["name"] 
                    date_of_birth = f"{officer['date_of_birth']['month']}/{officer['date_of_birth']['year']}"
                    
                    directors["name"] = name
                    directors["DoB"] = date_of_birth
                
            except (KeyError, TypeError): 
                print(f"Error parsing officer info on company: {company_number}")

        if directors: 
            return directors 
        else: 
            directors["name"] = "No director found"
            directors["DoB"] = "No director found"
            return directors
            




def get_charge_holder(company_number):

    # Request info on charges
    url = f"https://api.company-information.service.gov.uk/company/{company_number}/charges"
    response = requests.get(url, auth=HTTPBasicAuth(api_key, "")) 

    # Check server has loaded successfully
    if response.status_code == 200: 
        data = response.json()

        # Set variables to be used (using .get() so that if the key does not exist a default value is given instead of KeyError)
        total = data.get("total_count", 0) 
        satisfied = data.get("satisfied_count", 0)
        items = data.get("items", [])
        # Identify if there is a charge holder and output who it is if there is one
        if total == 0: 
            return {"status": "no charges", "charge_holders": None}
        elif total == satisfied: 
            return {"status": "all satisfied", "charge_holders": None}   
        
    # If there are outstanding charges
        charge_holders = []
        for item in items:
            if item.get("status") == "outstanding": 
                try:
                    holder_name = item["persons_entitled"][0]["name"] 
                    charge_holders.append(holder_name)

                except (KeyError, IndexError): 
                    print(f"Incomplete charge data for comapny {company_number}")
            
        return {"status": "outstanding", 
                "charge_holders": charge_holders if charge_holders else None}
            
         

def get_dormant_accounts(company_number): 
    
    # Request info on filing history
    url = f"https://api.company-information.service.gov.uk/company/{company_number}/filing-history?category=accounts"
    

    response = requests.get(url, auth=HTTPBasicAuth(api_key, "")) 

    if response.status_code == 200: 
        data = response.json()

        try:

            # Retrieve data on potential dormant accounts
            items = data.get("items", [])
            if len(items) == 0: 
                return "dissolved"
            description = items[0].get("description", None)
            # print(f"description: {description}")
            # Verify if accounts are dormant
            if description == None: 
                return "No description of accounts"
            
            if re.search(r"dormant", description.lower()):
                return "Yes" 
            
            return "No"
                
        # Catch the error if company has incomplete filing history
        except (TypeError, KeyError, IndexError):
            print(f"Error processing filing history for company {company_number}")


def get_overdue(company_number): 
    
    # Request info on company
    url = f"https://api.company-information.service.gov.uk/company/{company_number}"
    response = requests.get(url, auth=HTTPBasicAuth(api_key, "")) 

    # Verify page has loaded and extract json
    if response.status_code == 200: 
        data = response.json() 

        # Extract required info on accounts and confirmation statement from json and anticpate potential error for dissovled companies
        accounts = data.get("accounts", {})
        next_accounts = accounts.get("next_accounts", {})
        accounts_overdue = next_accounts.get("overdue")

        confirmation_statement = data.get("confirmation_statement", {}) 
        confirmation_statement_overdue = confirmation_statement.get("overdue", None)

        overdue = {}
        if accounts_overdue is True: 
            overdue["accounts"] = "overdue"
        elif accounts_overdue is False: 
            overdue["accounts"] = "up to date"
        else: 
            overdue["accounts"] = "dissolved"
        
        if confirmation_statement_overdue is True:  
            overdue["confirmation statement"] = "overdue"
        elif confirmation_statement_overdue is False: 
                overdue["confirmation statement"] = "up to date"
        else: 
            overdue["confirmation statement"] = "dissolved"
        
        return overdue

        
# def get_linked_in(name, company_name): 
#     query = "site:linkedin.com/in {name} Kent"
#     results = search(query, num_results=5)
#     for result in results: 
#         print(result)

if __name__ == "__main__":
    main()
# print(get_directors('OC446643')["name"])
# print(get_directors('OC446643')["DoB"])




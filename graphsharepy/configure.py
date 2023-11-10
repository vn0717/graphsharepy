"""  
Script to configure graphsharepy with the secrets needed to access SharePoint by Microsoft Graph

Michael P. Vossen
Created 9/22/2023

"""

import graphsharepy as gsp   
import getpass
import requests
import os
from shutil import copyfile
import secure



######################################################################################################
#The follwing area is for user inputs.  Exception handeling is includes to try and avoid later issues.
######################################################################################################

user = input("Enter Office 365 Email: ")

if user.find("@") == -1:
    raise ValueError(f"{user} is not a valid email.  Be sure to include the entire email address")

pas = getpass.getpass("Enter Office 365 Password: ")
host = input("Enter Full SharePoint Hostname (e.g., contoso.sharepoint.com): ")

if host.find(".sharepoint.com") == -1:
    raise ValueError(f"{host} is not a valid host name.  Host names are structured to be <your_company>.sharepoint.com (e.g., contoso.sharepoint.com)")


sharepoint = input("Enter SharePoint Name: ")
sharepoint = sharepoint.replace(" ", "")

exist = os.path.exists("secret.py")

if exist == True:
    file = open("secret.py", 'r')
    old_data = file.readlines()
    file.close()

user = user.strip()
pas = pas.strip()
host = host.strip()

sharepoint_exist = False
if exist == True:
    for line in old_data:
        if line.find(f"{sharepoint}_{user}_{host}") != -1:
            sharepoint_exist = True

    #wipe memory
    num = len(old_data)

    for i in range(num):
        for key in ['password', 'tenant', 'app_id', 'sec_val']:
            old_data[i] = secure.wipe_subval(key, old_data[i]) 

if sharepoint_exist == True:
    option = ""
else:
    option = "1"



print("\n\nFor the follwing prompts you must have completed the Azure app registration for SharePoint\n\n")
first = True
while option != "1" and option != "2":
    if first == True:
        option = input("\nThis registration already exists.  Is it a completly new app registration or a renewal?\n1: New App Registration | 2: Renewal\n")
        first = False
    else:
        option = input(f"\n{option} is an invalid option.  Is this a completly new app registration or a renewal?\n1: New App Registration | 2: Renewal\n")
    option = option.strip()



if option == "1":
    app_id = getpass.getpass("Enter Azure App Application ID: ")
sec_val = getpass.getpass("Enter Azure App Secret Value: ")








###########################################
# Find info that we can automatically find.
###########################################

if option == "1":

    end = host.find(".")
    prefix = host[:end]
    response = requests.get(f"https://login.microsoftonline.com/{prefix}.onmicrosoft.com/.well-known/openid-configuration").json()
    endpoint = response['token_endpoint']
    start = endpoint.find(".com") + 5
    end = endpoint.find("/oauth2")
    tenant_id = endpoint[start:end]
    





########################
# Write info to a module
########################


def replace_value(information, index_name, new_value):
    print(index_name)
    start = information.find(index_name)
    print(start)
    failure = False
    if start == -1:
        for key in ['password', 'tenant', 'app_id', 'sec_val']:
            information = secure.wipe_subval(key, information)
            failure = True 
    else:
        start_looking = False
        for starting_point in range(start, len(information)):
            if information[starting_point] == ":":
                start_looking = True
            if start_looking == True and information[starting_point] == "\'":
                starting_point += 1
                break
        for ending_point in range(starting_point, len(information)):
            if information[ending_point] == "\'":
                break
        information = f"{information[:starting_point]}{new_value}{information[ending_point:]}"
    new_value = secure.wipe_mem(new_value)
    return information, failure

if exist == True:
    file = open("secret.py", 'r')
    old_data = file.readlines()
    file.close()
    #eliminate dictornay stuff
    old_data = old_data[1:-1]
else:
    old_data = []

if option == "1" and sharepoint_exist == True:
    rid_lines = []
    for count, line in enumerate(old_data):
        if line.find(f"{sharepoint}_{user}_{host}") != -1:
            rid_lines.append(count)


    for adjust, line in enumerate(rid_lines):
        old_data.pop(line - adjust)

if option == "2":
    for count, line in enumerate(old_data):
        if line.find(f"{sharepoint}_{user}_{host}") != -1:
            break
    
    for variable, value in zip(["password", "sec_val"], [pas,sec_val]):
        old_data[count], fail_check = replace_value(old_data[count], variable, value)
        if fail_check == True:
            break
    
    if fail_check == True:
        pas = secure.wipe_mem(pas)
        sec_val = secure.wipe_mem(sec_val)  
        num = len(old_data)

        for i in range(num):
            for key in ['password', 'tenant', 'app_id', 'sec_val']:
                old_data[i] = secure.wipe_subval(key, old_data[i])
                
        raise Exception("There was an error in the configuration")

if option == "1":    
    old_data.append(f"\'{sharepoint}_{user}_{host}\'" + " : {" + f"'user':'{user}', 'password':'{pas}', 'host':'{host}', 'tenant':'{tenant_id}', 'app_id':'{app_id}', 'sec_val':'{sec_val}', 'sharepoint':'{sharepoint}'" + "},\n")

file = open("temp_secret.py", 'w')
file.write("secret_info = {\n")
file.writelines(old_data)
file.write("}")
    
file.close()


###################
# Secure the memory
###################

pas = secure.wipe_mem(pas)
sec_val = secure.wipe_mem(sec_val)  
num = len(old_data)

for i in range(num):
    for key in ['password', 'tenant', 'app_id', 'sec_val']:
        old_data[i] = secure.wipe_subval(key, old_data[i]) 
        
        
#save module to the correct name
if exist:
    os.remove("secret.py")
copyfile("temp_secret.py", "secret.py")
os.remove("temp_secret.py")

if option == "1":
    #run the initalization
    gsp.OAuth2_SharePoint(sharepoint, host, user, First=True)

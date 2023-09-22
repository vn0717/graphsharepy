"""  
Script to configure graphsharepy with the secrets needed to access SharePoint by Microsoft Graph

Michael P. Vossen
Created 9/22/2023

"""

#import graphsharepy as gsp   
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
host = input("Enter SharePoint Hostname: ")

if host.find(".sharepoint.com") == -1:
    raise ValueError(f"{host} is not a valid host name.  Host names are structured to be <your_company>.sharepoint.com (e.g., contoso.sharepoint.com)")

sharepoint = input("Enter SharePoint Name: ")

sharepoint = sharepoint.replace(" ", "")

print("\n\nFor the follwing prompts you must have completed the Azure app registration for SharePoint\n\n")
app_id = getpass.getpass("Enter Azure App Application ID: ")
sec_val = getpass.getpass("Enter Azure App Secret Value: ")

first = ""

while first.lower().find("y") == -1 and first.lower().find("n") == -1:
    first = input("Is this the first time this application is being ran? (y/n): ")





###########################################
# Find info that we can automatically find.
###########################################

end = host.find(".")
prefix = host[:end]
response = requests.get(f"https://login.microsoftonline.com/{prefix}.onmicrosoft.com/.well-known/openid-configuration").json()

print(response)
endpoint = response['token_endpoint']
start = endpoint.find(".com") + 5
end = endpoint.find("/oauth2")
tenant_id = endpoint[start:end]





########################
# Write info to a module
########################


exist = os.path.exists("secret.py")

if exist == True:
    file = open("secret.py", 'r')
    old_data = file.readlines()
    file.close()

else:
    old_data = []

rid_lines = []
for count, line in enumerate(old_data):
    if line.find(sharepoint) != -1:
        rid_lines.append(count)


for adjust, line in enumerate(rid_lines):
    old_data.pop(line - adjust)

file = open("temp_secret.py", 'w')

    
old_data.append(f"{sharepoint}" + " = {" + f"'user':'{user}', 'password':'{pas}', 'host':'{host}', 'tenant':'{tenant_id}', 'app_id':'{app_id}', 'sec_val':'{sec_val}', 'sharepoint':'{sharepoint}'" + "}")

file.writelines(old_data)
    
file.close()


###################
# Secure the memory
###################

pas = secure.wipe_mem(pas)

num = len(old_data)

for i in range(num):
    for key in ['password', 'tenant', 'app_id', 'sec_val']:
        old_data[i] = secure.wipe_subval(key, old_data[i]) 
        
        
#save module to the correct name
if exist:
    os.remove("secret.py")
copyfile("temp_secret.py", "secret.py")
os.remove("temp_secret.py")     
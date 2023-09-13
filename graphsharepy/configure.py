#import graphsharepy as gsp   
import getpass
import requests
import os

def wipe_mem(value):
    length = len(value)
    value = "0"
    for i in range(length-1):
        value += "0"
    return value

def wipe_subval(key, string):
    loc = string.find(key)
    if loc != -1:
        length = len(key)+1
        loc += length
        sub_string = string[loc:]
        end = sub_string.find(",")
        word = sub_string[:end]
        length_word = len(word)
        for i in range(length_word):
            word = word[:i] + "0" + word[i + 1:]
            string = string[:i+loc] + "0" + string[i+loc + 1:]

    return string
        


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
sec_id = getpass.getpass("Enter Azure App Secret ID: ")
sec_val = getpass.getpass("Enter Azure App Secret Value: ")



end = host.find(".")
prefix = host[:end]
response = requests.get(f"https://login.microsoftonline.com/{prefix}.onmicrosoft.com/.well-known/openid-configuration").json()
endpoint = response['token_endpoint']
start = endpoint.find(".com") + 5
end = endpoint.find("/oauth2")
tenant_id = endpoint[start:end]

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

    
old_data.append(f"{sharepoint}" + " = {" + f"'user':{user}, 'password':{pas}, 'host':{host}, 'tenant':{tenant_id}, 'sec_id':{sec_id}, 'sec_val':{sec_val}, 'sharepoint':{sharepoint}" + "}")

file.writelines(old_data)
    
file.close()


pas = wipe_mem(pas)

num = len(old_data)

for i in range(num):
    for key in ['password', 'tenant', 'sec_id', 'sec_val']:
        old_data[i] = wipe_subval(key, old_data[i]) 
        
print(old_data)       
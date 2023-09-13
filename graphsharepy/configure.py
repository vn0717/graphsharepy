#import graphsharepy as gsp   
import getpass
import requests

user = input("Enter Office 365 Email: ")

if user.find("@") == -1:
    raise ValueError(f"{user} is not a valid email.  Be sure to include the entire email address")

pas = getpass.getpass("Enter Office 365 Password: ")
host = input ("Enter SharePoint Hostname: ")

if host.find(".sharepoint.com") == -1:
    raise ValueError(f"{host} is not a valid host name.  Host names are structured to be <your_company>.sharepoint.com (e.g., contoso.sharepoint.com)")

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


file = open("secret.py", 'w')

for val, name in zip([user,pas,host,tenant_id, sec_id, sec_val], ["user", "password", "host", "tenant", "sec_id", "sec_val"]):
    file.write(f"{name} = '{val}'\n")
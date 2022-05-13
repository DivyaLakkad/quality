import requests

from getpass import getpass
from requests_ntlm import HttpNtlmAuth

url = "https://grahamcanada.sharepoint.com/sites/Quinn/Documents/Quality_test/In_test/Book.xlsx"

session = requests.Session()
session.verify = False

username = input("basharatj@graham.ca")
password = getpass("grahamindustrial")

session.auth = HttpNtlmAuth(username, password)
response = session.get(url)

with open(sharepointoutput.xlsx, wb) as f:
    f.write(response.content)
import requests
import pandas as pd
from webbrowser import open as open_url
import sys

# Asking if want to use admin API or not
admin = ""
while admin != "Y" and admin != "N" and admin != "X":
    admin = input("Run APIs as admin (Y/N), or enter X to exit: ").upper()
    if admin == "X":
        sys.exit(1)


# Getting the bearer token
print("Copy the bearer token from the Request Preview section of the below website. Only copy the string after Bearer word.")
if (admin == "Y"):
    toekn_url = "https://learn.microsoft.com/en-us/rest/api/power-bi/admin/reports-get-reports-as-admin#code-try-0"
else:
    toekn_url = "https://learn.microsoft.com/en-us/rest/api/power-bi/reports/get-reports#code-try-0"

try:
    print("Trying to automatically open the website")
    open_url(toekn_url, new=0, autoraise=True)
except Exception as e:
    print("Unsuccessful in opening the website. Please open the below url")
    print(toekn_url)

# Setting up the authorization header
access_token = input("Enter Your bearer token: ")
headers = {
    "Authorization": "Bearer {}.format(access_token)"
}


# Get all workspaces
print("Getting all workspaces")
if (admin == "Y"):
    top = 1000
    # Getting the top 1000 workspaces. Max limit of api is 5000
    workspaces_url = "https://api.powerbi.com/v1.0/myorg/admin/groups?$top={}".format(
        top)
else:
    workspaces_url = "https://api.powerbi.com/v1.0/myorg/groups"
workspaces_response = requests.get(workspaces_url, headers=headers)
if 400 <= workspaces_response.status_code < 500:
    print("Error getting data. Status code: ", workspaces_response.status_code)
    sys.exit(1)
workspaces_data = workspaces_response.json()
print("Found {} workspaces".format(len(workspaces_data["value"])))
workspaces_df = pd.DataFrame(workspaces_data["value"])

# Get all  Reports
print("Getting all reports")
if (admin == "Y"):
    reports_url = "https://api.powerbi.com/v1.0/myorg/admin/reports"
else:
    reports_url = "https://api.powerbi.com/v1.0/myorg/reports"
reports_response = requests.get(reports_url, headers=headers)
if 400 <= reports_response.status_code < 500:
    print("Error getting data. Status code: ", reports_response.status_code)
    sys.exit(1)
reports_data = reports_response.json()
print("Found {} reports".format(len(reports_data["value"])))
reports_df = pd.DataFrame(reports_data["value"])


# Get Datasources
print("Getting Data sources of the report")
datasource_df = pd.DataFrame()

for i in range(0, reports_df.shape[0]):
    print("Getting for {} index report.".format(i))

    datasetId = reports_df.iloc[i]["datasetId"]

    if (admin == "Y"):
        datasource_url = "https://api.powerbi.com/v1.0/myorg/admin/datasets/{}/datasources".format(
            datasetId)
    else:
        datasource_url = "https://api.powerbi.com/v1.0/myorg/datasets/{}/datasources".format(
            datasetId)

    datasource_response = requests.get(datasource_url, headers=headers)
    datasource_data = ""
    print(datasource_url)
    if 400 <= datasource_response.status_code < 500:
        if (datasource_response.status_code == 404):
            datasource_data = {}
        else:
            print("Error getting data. Status code: ",
                  datasource_response.status_code)
            sys.exit(1)
    if (datasource_data == {}):
        pass
    else:
        datasource_data = datasource_response.json()
    temp_datasource_df = pd.DataFrame()
    try:
        value = datasource_data["value"]
        if (value == []):
            temp_datasource_df.at[0, "error"] = False
            temp_datasource_df.at[0, "reportName"] = reports_df.iloc[i]["name"]
            temp_datasource_df.at[0, "reportId"] = reports_df.iloc[i]["id"]
            temp_datasource_df.at[0, "datasetId"] = datasetId
            temp_datasource_df.at[0, "rawResonse"] = str(datasource_data)
            temp_datasource_df.at[0, "data"] = False
            temp_datasource_df.at[0, "value"] = "[]"

        else:
            for n in range(0, len(value)):
                temp_datasource_df.at[0, "error"] = False
                temp_datasource_df.at[0,
                                      "reportName"] = reports_df.iloc[i]["name"]
                temp_datasource_df.at[0, "reportId"] = reports_df.iloc[i]["id"]
                temp_datasource_df.at[0, "datasetId"] = datasetId
                temp_datasource_df.at[0, "rawResonse"] = str(datasource_data)
                temp_datasource_df.at[0, "data"] = True
                temp_datasource_df.at[0, "value"] = str(value)

                for key in value[n].keys():
                    temp_datasource_df.at[0, key] = str(value[n][key])

                    if (key == 'connectionDetails'):
                        for k in value[n][key].keys():
                            temp_datasource_df.at[0, k] = str(value[n][key][k])

        datasource_df = pd.concat(
            [datasource_df, temp_datasource_df], ignore_index=True)

    except Exception as e:
        temp_datasource_df.at[0, "error"] = True
        temp_datasource_df.at[0, "reportName"] = reports_df.iloc[i]["name"]
        temp_datasource_df.at[0, "reportId"] = reports_df.iloc[i]["id"]
        temp_datasource_df.at[0, "datasetId"] = datasetId
        temp_datasource_df.at[0, "rawResonse"] = str(datasource_data)
        temp_datasource_df.at[0, "data"] = False
        temp_datasource_df.at[0, "value"] = "[]"
        temp_datasource_df.at[0, "errorReason"] = str(e)
        datasource_df = pd.concat(
            [datasource_df, temp_datasource_df], ignore_index=True)


# Writing all the data into an excel file
print("Writing all the data into an excel file")
with pd.ExcelWriter('powerBI_datasets.xlsx') as writer:
    workspaces_df.to_excel(writer, sheet_name='workspaces')
    reports_df.to_excel(writer, sheet_name='reports')
    datasource_df.to_excel(writer, sheet_name='datasources')
print("Suceessfully Written. Excel file with the name powerBI_datasets.xlsx is created in the same location where this code is located")

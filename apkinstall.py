import requests
import time
import pandas as pd
import urllib3

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

BASE_URL = ""
OPS_RAMP_SECRET = ''
OPS_RAMP_KEY = ''
EXCEL_FILE_PATH = "install_apps.xlsx"

def get_access_token(retries=3, delay=5):
    token_url = BASE_URL + "auth/oauth/token"
    auth_data = {
        'client_secret': OPS_RAMP_SECRET,
        'grant_type': 'client_credentials',
        'client_id': OPS_RAMP_KEY
    }
    headers = {'Content-Type': 'application/x-www-form-urlencoded'}
    
    for i in range(retries):
        try:
            token_response = requests.post(token_url, data=auth_data, headers=headers, verify=False)
            token_response.raise_for_status()
            token_data = token_response.json()
            if "invalid_token" in str(token_data):
                print("Received invalid token, retrying...")
                continue  # Retry if the token is invalid
            return token_data.get('access_token')
        except requests.exceptions.RequestException as e:
            print(f"Attempt {i + 1} failed: {str(e)}")
            if i < retries - 1:
                time.sleep(delay)
            else:
                print("Max retries reached. Failed to generate the token.")
                return None

def install_app(access_token, tenant_id, tenant_name, app_name, app_name_caps_):
    auth_header = {'Authorization': 'Bearer ' + access_token, 'Content-Type': 'application/json'}
    api_endpoint = f'https://www/api/v2/tenants/{tenant_id}/integrations/install/{app_name_caps_}'
    payload = {
        "displayName": app_name,
        "category": "REPORTING_APPS"
    }
    
    try:
        response = requests.post(api_endpoint, headers=auth_header, json=payload, verify=False)
        response.raise_for_status()
        print(f"{app_name} installed successfully for client {tenant_name}!")
    except requests.exceptions.RequestException as e:
        if response.status_code == 500:
            if "already exists" in response.text:
                print(f"{app_name} is already installed for client {tenant_name}.")
            else:
                print(f"Failed to install {app_name} for client {tenant_name}. Status code: {response.status_code}")
                print("Error message:", response.text)
        else:
            print(f"Failed to install {app_name} for client {tenant_name}. Status code: {response.status_code}")
            print("Error message:", response.text)

def main():
    access_token = get_access_token()
    if not access_token:
        return

    try:
        df = pd.read_excel(EXCEL_FILE_PATH)
    except Exception as e:
        print(f"Error reading the Excel file: {str(e)}")
        return

    # List of apps to be installed
    apps_list = ["App names"]

    for index, row in df.iterrows():
        client_id = row['Client_ID']
        tenant_name = row['Tenant_Name']
        
        for app_name in apps_list:
            app_name_caps = app_name.upper().replace(" ", "-")
            install_app(access_token, client_id, tenant_name, app_name, app_name_caps)

if __name__ == "__main__":
    main()

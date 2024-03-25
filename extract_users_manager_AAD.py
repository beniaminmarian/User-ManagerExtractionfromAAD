import requests
import pandas as pd

# Azure AD App Registration details
tenant_id = '6df5d9b1-2807-4c12-a223-63909d98a6f2'
client_id = 'e2f31f43-9c6d-477c-9702-5ebdac5004ba'
client_secret = 'NaM8Q~Rp_nBTZr0gW6cMgJYS5i4m-Mgsn4_IYc5w'
scope = 'https://graph.microsoft.com/.default'
token_endpoint = f'https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token'

# Request an access token
token_data = {
    'grant_type': 'client_credentials',
    'client_id': client_id,
    'client_secret': client_secret,
    'scope': scope
}

token_response = requests.post(token_endpoint, data=token_data)
access_token = token_response.json().get('access_token')

# Use the access token to make requests to Microsoft Graph API
graph_users_endpoint = 'https://graph.microsoft.com/v1.0/users?$select=id,displayName,jobTitle,userType,accountEnabled'

# Initialize an empty list to store user data
all_users = []
i=1

# Make a request to get all users
while graph_users_endpoint:
    response = requests.get(graph_users_endpoint, headers={'Authorization': f'Bearer {access_token}'})
    users = response.json().get('value', [])
    all_users.extend(users)
    graph_users_endpoint = response.json().get('@odata.nextLink')

# Create a list of dictionaries to store user and manager data
user_manager_list = []

# Process the response to extract display name and manager
for user in all_users:
    print('Processing ' + str(i) + '/' +str(len(all_users)))
    i=i+1
    display_name = user.get('displayName')
    job_title = user.get('jobTitle')
    user_type = user.get('userType')
    account_enabled = user.get('accountEnabled')
    manager_url = f"https://graph.microsoft.com/v1.0/users/{user['id']}/manager"
    manager_response = requests.get(manager_url, headers={'Authorization': f'Bearer {access_token}'})
    manager_data = manager_response.json()
    manager_name = manager_data.get('displayName')
  
# Handle the display name, replacing or omitting problematic characters
    if display_name:
        display_name = display_name.replace('\u0103', '')  # Replace the problematic character

    # Handle the manager name, replacing or omitting problematic characters
    if manager_name:
        manager_name = manager_name.replace('\u0103', '')  # Replace the problematic character

    # Add user and manager data to the list
    user_manager_list.append({'Name': display_name, 'Manager': manager_name, 'Job Title': job_title, 'User Type': user_type, 'Account Enabled': account_enabled})
    

# Create a DataFrame from the list
df = pd.DataFrame(user_manager_list)


# Export the DataFrame to a CSV file
df.to_csv("tecknoworkers-new.csv", index=False, encoding='utf-8-sig')

# Print a success message
print("Entries added to DataFrame and exported to tecknoworkers-new.csv")

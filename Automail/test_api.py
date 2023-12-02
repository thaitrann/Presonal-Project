import requests
import json

client_id = 'ac6f6e89-959e-4dc7-a936-82600227c88c'
client_secret = 'hNt8Q~GbFwYP6wOU~GmBnG1low6dMF2Hch-tUaJ_'
tenant_id = 'b0cfcd2e-b960-4760-9f29-664a4c369b4a'
grant_type = 'client_credentials'
resdirect_uri = 'https://localhost'
scope = 'https://graph.microsoft.com/.default'
refresh_token = 'refresh_token'

token_url = f'https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token'

token_data = {
    'client_id': client_id,
    'client_secret': client_secret,
    'grant_type': grant_type,
    'resdirect_uri': resdirect_uri,
    'scope': scope,
    'refresh_token': refresh_token
}

token_response = requests.post(token_url, data=token_data)
access_token = json.loads(token_response.text)['access_token']

header = {
    'Authorization': f'Bearer {access_token}',
    'Content-Type': 'application/json'
}

folder_id_url = f'https://graph.microsoft.com/v1.0/{client_id}/drive/root/children'

folder_id_respone = requests.get(folder_id_url, headers=header)

if folder_id_respone.status_code == 200:
    result = folder_id_respone.json()
    print(result)
else:
    print(f'Error: {folder_id_respone.text}')


import requests, json
API_KEY="4851c5dc2cad4fdc996da5a347965c57"
BASE_URL="https://apim-api.noga-iso.co.il/productionmix/PRODMIXAPI/v1"
headers={"Ocp-Apim-Subscription-Key": API_KEY, "Content-Type":"application/json"}
payload={"fromDate":"26-05-2024","toDate":"26-05-2024"}
resp=requests.post(BASE_URL, headers=headers, json=payload)
print("Status:", resp.status_code)
print("Content-Type:", resp.headers.get("content-type"))
print(resp.text[:1000]) 
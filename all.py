import requests

# Store API details
store_id = "111111..."
api_token = "XXXXXXXXXXX..."

# API endpoint to fetch orders
url = f"https://app.ecwid.com/api/v3/{store_id}/orders"
headers = {
    "Authorization": f"Bearer {api_token}"
}

# Fetch orders
response = requests.get(url, headers=headers)

if response.status_code == 200:
    orders = response.json().get("items", [])
    # Print the raw order data for the first order
    if orders:
        print(orders[0])  # Print the first order for inspection
else:
    print("Failed to retrieve data. Status Code:", response.status_code)

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

def print_order_details(order, indent=0):
    for key, value in order.items():
        if isinstance(value, dict):  # Check if the value is a dictionary
            print("  " * indent + f"{key}:")
            print_order_details(value, indent + 1)  # Recursively print nested dicts
        elif isinstance(value, list):  # Check if the value is a list
            print("  " * indent + f"{key}:")
            for item in value:  # Iterate through list items
                if isinstance(item, dict):
                    print_order_details(item, indent + 1)  # Print each item details
                else:
                    print("  " * (indent + 1) + f"- {item}")  # Print simple list item
        else:
            print("  " * indent + f"{key}: {value}")

if response.status_code == 200:
    orders = response.json().get("items", [])
    if orders:
        first_order = orders[12]
        print_order_details(first_order)
else:
    print("Failed to retrieve data. Status Code:", response.status_code)

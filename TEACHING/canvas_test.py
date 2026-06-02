import requests

TOKEN = "YOUR_TOKEN_HERE"
BASE_URL = "https://canvas.spcollege.edu/api/v1"

headers = {"Authorization": f"Bearer {TOKEN}"}

response = requests.get(f"{BASE_URL}/courses/32441", headers=headers)
print(response.status_code)
print(response.json())

import requests
from bs4 import BeautifulSoup

TOKEN = "YOUR_TOKEN_HERE"
BASE = "https://canvas.spcollege.edu/api/v1"
COURSE = 32441

headers = {"Authorization": f"Bearer {TOKEN}"}

for assignment_id in [751559, 751560]:
    r = requests.get(f"{BASE}/courses/{COURSE}/assignments/{assignment_id}", headers=headers)
    a = r.json()
    print(f"\n=== Assignment {assignment_id}: {a['name']} ===")
    soup = BeautifulSoup(a.get("description", ""), "html.parser")
    for iframe in soup.find_all("iframe"):
        src = iframe.get("src", "")
        if "youtube" in src or "youtu.be" in src:
            print("YouTube iframe attributes:")
            for attr, val in iframe.attrs.items():
                print(f"  {attr}: {val}")

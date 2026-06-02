import requests

TOKEN = "YOUR_TOKEN_HERE"
BASE = "https://canvas.spcollege.edu/api/v1"
COURSE = 32441

headers = {"Authorization": f"Bearer {TOKEN}"}

for assignment_id in [751559, 751560]:
    r = requests.get(f"{BASE}/courses/{COURSE}/assignments/{assignment_id}", headers=headers)
    a = r.json()
    print(f"\n=== Assignment {assignment_id}: {a['name']} ===")
    print(a.get("description", ""))
    print("--- END ---")

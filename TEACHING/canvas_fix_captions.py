import requests

TOKEN = "YOUR_TOKEN_HERE"
BASE = "https://canvas.spcollege.edu/api/v1"
COURSE = 32441
ASSIGNMENT_ID = 751560

headers = {"Authorization": f"Bearer {TOKEN}"}

# Fetch current description
r = requests.get(f"{BASE}/courses/{COURSE}/assignments/{ASSIGNMENT_ID}", headers=headers)
a = r.json()
html = a.get("description", "")

# Add cc_load_policy=1 to the YouTube embed URL
old_src = "https://www.youtube.com/embed/D1Na7fbpVdo?si=S6XevkUOyQgI5lbQ"
new_src = "https://www.youtube.com/embed/D1Na7fbpVdo?si=S6XevkUOyQgI5lbQ&cc_load_policy=1"

new_html = html.replace(old_src, new_src)

if new_html == html:
    print("URL not found in assignment - no changes made.")
else:
    r2 = requests.put(
        f"{BASE}/courses/{COURSE}/assignments/{ASSIGNMENT_ID}",
        headers=headers,
        json={"assignment": {"description": new_html}}
    )
    print(f"Status: {r2.status_code}")
    print("Done - check A02 in Canvas and run the accessibility checker again.")

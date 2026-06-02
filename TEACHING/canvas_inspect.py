import requests
from bs4 import BeautifulSoup

TOKEN = "YOUR_TOKEN_HERE"
BASE = "https://canvas.spcollege.edu/api/v1"
COURSE = 32441

headers = {"Authorization": f"Bearer {TOKEN}"}

def get_all(url, params=None):
    results = []
    while url:
        r = requests.get(url, headers=headers, params=params)
        results.extend(r.json())
        url = r.links.get("next", {}).get("url")
        params = None
    return results

def find_youtube_iframes(html, source):
    soup = BeautifulSoup(html or "", "html.parser")
    for iframe in soup.find_all("iframe"):
        src = iframe.get("src", "")
        if "youtube" in src or "youtu.be" in src:
            print(f"\n--- Found in: {source} ---")
            print(iframe)

# Pages
print("=== PAGES ===")
pages = get_all(f"{BASE}/courses/{COURSE}/pages", {"per_page": 100})
for page in pages:
    detail = requests.get(f"{BASE}/courses/{COURSE}/pages/{page['url']}", headers=headers).json()
    find_youtube_iframes(detail.get("body", ""), f"Page: {page['title']}")

# Assignments
print("\n=== ASSIGNMENTS ===")
assignments = get_all(f"{BASE}/courses/{COURSE}/assignments", {"per_page": 100})
for a in assignments:
    find_youtube_iframes(a.get("description", ""), f"Assignment: {a['name']}")

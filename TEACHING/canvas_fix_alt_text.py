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

def fix_images(html):
    if not html:
        return html, False
    soup = BeautifulSoup(html, "html.parser")
    changed = False
    for img in soup.find_all("img"):
        alt = img.get("alt", "").strip()
        # Fix if blank or looks like a filename
        if not alt or alt.lower().endswith((".png", ".jpg", ".jpeg", ".gif", ".webp")):
            img["alt"] = "SolidWorks tutorial image"
            changed = True
    return str(soup), changed

# Fix Pages
print("=== PAGES ===")
pages = get_all(f"{BASE}/courses/{COURSE}/pages", {"per_page": 100})
for page in pages:
    detail = requests.get(f"{BASE}/courses/{COURSE}/pages/{page['url']}", headers=headers).json()
    new_html, changed = fix_images(detail.get("body", ""))
    if changed:
        r = requests.put(f"{BASE}/courses/{COURSE}/pages/{page['url']}", headers=headers,
                         json={"wiki_page": {"body": new_html}})
        print(f"Updated page: {page['title']} ({r.status_code})")

# Fix Assignments
print("\n=== ASSIGNMENTS ===")
assignments = get_all(f"{BASE}/courses/{COURSE}/assignments", {"per_page": 100})
for a in assignments:
    new_html, changed = fix_images(a.get("description", ""))
    if changed:
        r = requests.put(f"{BASE}/courses/{COURSE}/assignments/{a['id']}", headers=headers,
                         json={"assignment": {"description": new_html}})
        print(f"Updated assignment: {a['name']} ({r.status_code})")

print("\nDone.")

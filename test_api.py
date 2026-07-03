import json
import sys
from api.index import app

def test_scrape_endpoint():
    print("Testing /api/index with action: scrape...")
    client = app.test_client()
    
    payload = {
        "action": "scrape",
        "urls": ["https://example.com"]
    }
    
    response = client.post('/api/index', json=payload)
    print(f"Status Code: {response.status_code}")
    
    if response.status_code == 200:
        data = json.loads(response.data)
        print("Response Data:", json.dumps(data, indent=2))
        if "results" in data and len(data["results"]) > 0:
            print("SUCCESS: Scrape endpoint is working!")
        else:
            print("FAILED: Results array is empty.")
            sys.exit(1)
    else:
        print(f"FAILED: Expected 200 OK, got {response.status_code}")
        print("Response:", response.data.decode('utf-8'))
        sys.exit(1)

if __name__ == "__main__":
    test_scrape_endpoint()

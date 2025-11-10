#!/usr/bin/env python3
"""
Test if Apps Script URL is accessible
"""

import requests
import json

url = "https://script.google.com/a/macros/commerceiq.ai/s/AKfycbyhw1FExly5FhiJCZEgyBcHZgSQesHdlJEFOX2GLfQaHOYG67YuiBt0CqRfOQEVz1QZ/exec"

print("Testing Apps Script URL...")
print(f"URL: {url}\n")

# Test 1: Simple GET request
print("1. Testing GET request...")
try:
    response = requests.get(url, timeout=10)
    print(f"   Status: {response.status_code}")
    print(f"   Content-Type: {response.headers.get('content-type', 'unknown')}")
    print(f"   Response length: {len(response.content)} bytes")
    if response.status_code == 200:
        print(f"   Response preview: {response.text[:200]}")
    else:
        print(f"   Response preview: {response.text[:500]}")
except Exception as e:
    print(f"   Error: {str(e)}")

print("\n" + "="*60 + "\n")

# Test 2: POST with minimal JSON
print("2. Testing POST request with minimal JSON...")
try:
    test_payload = {
        'data': [['Header1', 'Header2'], ['Value1', 'Value2']],
        'filename': 'test.xlsx',
        'sheet_name': 'Bob Perf Report'
    }
    
    response = requests.post(
        url,
        json=test_payload,
        headers={'Content-Type': 'application/json'},
        timeout=30
    )
    
    print(f"   Status: {response.status_code}")
    print(f"   Content-Type: {response.headers.get('content-type', 'unknown')}")
    print(f"   Response length: {len(response.content)} bytes")
    
    if response.status_code == 200:
        try:
            result = response.json()
            print(f"   ✅ SUCCESS!")
            print(f"   Response: {json.dumps(result, indent=2)}")
        except:
            print(f"   Response text: {response.text[:500]}")
    else:
        print(f"   ❌ FAILED")
        print(f"   Response text: {response.text[:1000]}")
        
except Exception as e:
    print(f"   Error: {str(e)}")
    import traceback
    traceback.print_exc()

print("\n" + "="*60)
print("\n💡 If you see 401/403 errors:")
print("   1. Go to https://script.google.com")
print("   2. Open your Apps Script project")
print("   3. Click 'Deploy' → 'Manage deployments'")
print("   4. Edit the deployment:")
print("      - Execute as: Me")
print("      - Who has access: Anyone (or 'Anyone in your organization')")
print("   5. Click 'Update'")
print("\n💡 If you see 404 errors:")
print("   - The deployment ID might be wrong")
print("   - Create a new deployment and update the URL")


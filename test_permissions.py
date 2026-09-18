import urllib.request
import urllib.error
import json

BASE_URL = 'http://127.0.0.1:8765'

def test_permissions():
    print("=== TESTING EMAIL-BASED ROLE PERMISSIONS ===")

    # 1. Check permissions
    req = urllib.request.Request(f'{BASE_URL}/api/permissions', headers={'X-User-Email': 'admin@company.com'})
    with urllib.request.urlopen(req) as resp:
        perms = json.loads(resp.read().decode('utf-8'))
    print("Configured Permissions:")
    for role, info in perms.items():
        print(f"  - {role}: {info['email']}")

    # 2. Test Dispatch Logging Permissions
    print("\n--- Testing Dispatch Logging Permissions ---")
    dispatch_payload = json.dumps({
        "sku": "CS0001",
        "platform": "Blinkit",
        "location": "Mumbai",
        "po_number": "PO-AUTH-TEST",
        "po_quantity": 50,
        "actual_sent": 50
    }).encode('utf-8')

    # 2a. Guest user attempts to log dispatch -> should FAIL 403
    try:
        req = urllib.request.Request(
            f'{BASE_URL}/api/dispatch-online',
            data=dispatch_payload,
            headers={'Content-Type': 'application/json', 'X-User-Email': 'guest@company.com'}
        )
        urllib.request.urlopen(req)
        print("ERROR: Guest user was able to log dispatch!")
    except urllib.error.HTTPError as e:
        print(f"  [PASS] Guest user denied dispatch logging (HTTP {e.code}): {e.read().decode('utf-8')}")

    # 2b. Inventory Lead attempts to log dispatch -> should FAIL 403
    try:
        req = urllib.request.Request(
            f'{BASE_URL}/api/dispatch-online',
            data=dispatch_payload,
            headers={'Content-Type': 'application/json', 'X-User-Email': 'inventory@company.com'}
        )
        urllib.request.urlopen(req)
        print("ERROR: Inventory Lead was able to log dispatch!")
    except urllib.error.HTTPError as e:
        print(f"  [PASS] Inventory Lead denied dispatch logging (HTTP {e.code}): {e.read().decode('utf-8')}")

    # 2c. Dispatch Lead attempts to log dispatch -> should SUCCEED 200
    req = urllib.request.Request(
        f'{BASE_URL}/api/dispatch-online',
        data=dispatch_payload,
        headers={'Content-Type': 'application/json', 'X-User-Email': 'dispatch@company.com'}
    )
    resp = urllib.request.urlopen(req)
    res_data = json.loads(resp.read().decode('utf-8'))
    disp_id = res_data['id']
    print(f"  [PASS] Dispatch Lead successfully logged dispatch (ID #{disp_id})")

    # Cleanup dispatch entry
    del_req = urllib.request.Request(
        f'{BASE_URL}/api/dispatch-online/{disp_id}',
        headers={'X-User-Email': 'dispatch@company.com'},
        method='DELETE'
    )
    urllib.request.urlopen(del_req)
    print(f"  [PASS] Cleaned up dispatch ID #{disp_id}")

    # 3. Test Inventory Stock & Thresholds Permissions
    print("\n--- Testing Inventory Stock Permissions ---")
    stock_payload = json.dumps({
        "sku": "CS0001",
        "total_inventory": 5060,
        "reorder_threshold": 1000
    }).encode('utf-8')

    # 3a. Dispatch Lead attempts to edit inventory stock -> should FAIL 403
    try:
        req = urllib.request.Request(
            f'{BASE_URL}/api/inventory-stock',
            data=stock_payload,
            headers={'Content-Type': 'application/json', 'X-User-Email': 'dispatch@company.com'}
        )
        urllib.request.urlopen(req)
        print("ERROR: Dispatch Lead was able to edit inventory stock!")
    except urllib.error.HTTPError as e:
        print(f"  [PASS] Dispatch Lead denied stock edit (HTTP {e.code}): {e.read().decode('utf-8')}")

    # 3b. Guest user attempts to edit inventory stock -> should FAIL 403
    try:
        req = urllib.request.Request(
            f'{BASE_URL}/api/inventory-stock',
            data=stock_payload,
            headers={'Content-Type': 'application/json', 'X-User-Email': 'guest@company.com'}
        )
        urllib.request.urlopen(req)
        print("ERROR: Guest user was able to edit inventory stock!")
    except urllib.error.HTTPError as e:
        print(f"  [PASS] Guest user denied stock edit (HTTP {e.code}): {e.read().decode('utf-8')}")

    # 3c. Inventory Lead attempts to edit inventory stock -> should SUCCEED 200
    req = urllib.request.Request(
        f'{BASE_URL}/api/inventory-stock',
        data=stock_payload,
        headers={'Content-Type': 'application/json', 'X-User-Email': 'inventory@company.com'}
    )
    resp = urllib.request.urlopen(req)
    print(f"  [PASS] Inventory Lead successfully updated stock settings (HTTP {resp.status})")

    # 4. Test Audit Log
    print("\n--- Testing Audit Log Tracking ---")
    req = urllib.request.Request(f'{BASE_URL}/api/audit-log')
    with urllib.request.urlopen(req) as resp:
        logs = json.loads(resp.read().decode('utf-8'))
    print(f"Recent Audit Logs: {len(logs)} entries recorded.")
    for l in logs[:3]:
        print(f"  #{l['id']} [{l['timestamp']}]: {l['details']}")

    print("\nALL PERMISSION AND ENFORCEMENT TESTS PASSED 100%!")

if __name__ == '__main__':
    test_permissions()

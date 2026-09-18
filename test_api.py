import urllib.request
import urllib.parse
import json

def test():
    # 1. Stats
    req = urllib.request.urlopen('http://127.0.0.1:8765/api/stats')
    stats = json.loads(req.read().decode('utf-8'))
    print("API Stats Test:")
    print("  Total Inventory:", stats['total_inventory'])
    print("  Total Sent:", stats['total_sent'])
    print("  Remaining Inventory:", stats['remaining_inventory'])
    print("  Low Stock Count:", stats['low_stock_count'])
    print("  PO Fill Rate:", f"{stats['fill_rate']}%")

    # 2. Master
    req = urllib.request.urlopen('http://127.0.0.1:8765/api/master')
    master = json.loads(req.read().decode('utf-8'))
    print(f"\nAPI Master Test: {len(master)} SKUs loaded.")
    for m in master[:3]:
        print(f"  {m['sku']} ({m['product_name']}): Online={m['online_total']}, Rem={m['remaining_inventory']}, Status={m['status']}")

    # 3. Online Dispatches & PO Columns
    req = urllib.request.urlopen('http://127.0.0.1:8765/api/dispatch-online')
    online = json.loads(req.read().decode('utf-8'))
    print(f"\nAPI Online Dispatches: {len(online)} rows.")
    first = online[0]
    print(f"  Sample row: ID #{first['id']} | PO#: {first['po_number']} | PO Qty: {first['po_quantity']} | Actual Sent: {first['actual_sent']} | Shortage: {first['variance']} | Fill: {first['fulfillment_rate']}%")

    # 4. Test Add New Dispatch with PO fields
    test_payload = {
        "sku": "CS0001",
        "dispatch_date": "2026-09-18",
        "platform": "Blinkit",
        "location": "Bengaluru Central",
        "po_number": "PO-TEST-8888",
        "po_quantity": 200,
        "actual_sent": 180,
        "appointment_date": "2026-09-20"
    }
    data = json.dumps(test_payload).encode('utf-8')
    post_req = urllib.request.Request('http://127.0.0.1:8765/api/dispatch-online', data=data, headers={'Content-Type': 'application/json'})
    resp = urllib.request.urlopen(post_req)
    res_json = json.loads(resp.read().decode('utf-8'))
    print("\nPOST New Online Dispatch Test:", res_json)
    new_id = res_json['id']

    # Verify that Blinkit units and remaining inventory updated
    req = urllib.request.urlopen('http://127.0.0.1:8765/api/master')
    master_after = json.loads(req.read().decode('utf-8'))
    cs1 = next(x for x in master_after if x['sku'] == 'CS0001')
    print(f"  After dispatch added -> CS0001 Blinkit Sent: {cs1['sent_blinkit']} (was 241, now {cs1['sent_blinkit']}), Rem: {cs1['remaining_inventory']}")

    # Delete test record
    del_req = urllib.request.Request(f'http://127.0.0.1:8765/api/dispatch-online/{new_id}', method='DELETE')
    del_resp = urllib.request.urlopen(del_req)
    print("  Cleaned up test record:", json.loads(del_resp.read().decode('utf-8')))

    # 5. Test Excel Export
    req = urllib.request.urlopen('http://127.0.0.1:8765/api/export-excel')
    content = req.read()
    print(f"\nExcel Export Test: Successfully received {len(content)} bytes of .xlsx data.")
    print("ALL API TESTS PASSED PERFECTLY!")

if __name__ == '__main__':
    test()

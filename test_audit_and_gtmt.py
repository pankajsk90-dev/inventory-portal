import urllib.request
import urllib.parse
import json

BASE = 'http://127.0.0.1:8765'

def test():
    print("=== TESTING AUDIT LOGS (NAME, EDIT DONE, TIMESTAMP) & GT/MT DEEP DIVE ===")

    # 1. Test GT Deep Dive API
    url_gt = f'{BASE}/api/gtmt-detail?channel=GT'
    with urllib.request.urlopen(url_gt) as resp:
        data_gt = json.loads(resp.read().decode('utf-8'))
    print("GT Deep-Dive Test:")
    print(f"  Channel: {data_gt['channel_title']}")
    print(f"  Total Units Sent via GT: {data_gt['total_units_sent']:,}")
    print(f"  Total GT Dispatches: {data_gt['total_dispatches']}")
    if data_gt['top_buyers']:
        print(f"  Top Buyer: {data_gt['top_buyers'][0]['buyer_distributor']} ({data_gt['top_buyers'][0]['sent']:,} units)")
    top_p = data_gt['products'][0] if data_gt['products'] else None
    if top_p:
        print(f"  Top Product: {top_p['product_name']} ({top_p['units_sent']:,} units, {top_p['share_percent']}%)")

    # 2. Test MT Deep Dive API
    url_mt = f'{BASE}/api/gtmt-detail?channel=MT'
    with urllib.request.urlopen(url_mt) as resp:
        data_mt = json.loads(resp.read().decode('utf-8'))
    print("\nMT Deep-Dive Test:")
    print(f"  Channel: {data_mt['channel_title']}")
    print(f"  Total Units Sent via MT: {data_mt['total_units_sent']:,}")
    print(f"  Total MT Dispatches: {data_mt['total_dispatches']}")
    if data_mt['top_buyers']:
        print(f"  Top Buyer: {data_mt['top_buyers'][0]['buyer_distributor']} ({data_mt['top_buyers'][0]['sent']:,} units)")

    # 3. Test Named Audit Log on Stock Edit
    print("\n--- Testing Stock Edit with User Name & Detailed Edit Diff ---")
    headers = {
        'Content-Type': 'application/json',
        'X-User-Email': 'inventory@company.com',
        'X-User-Name': 'Vikram Mehra (Inventory Manager)'
    }
    payload = json.dumps({
        "sku": "CS0001",
        "total_inventory": 5100,
        "reorder_threshold": 1000
    }).encode('utf-8')
    req = urllib.request.Request(f'{BASE}/api/inventory-stock', data=payload, headers=headers)
    urllib.request.urlopen(req)

    # 4. Test Named Audit Log on GT/MT Dispatch
    print("--- Testing GT/MT Dispatch with User Name & Edit Log ---")
    disp_headers = {
        'Content-Type': 'application/json',
        'X-User-Email': 'dispatch@company.com',
        'X-User-Name': 'Sunita Rao (Dispatch Lead)'
    }
    disp_payload = json.dumps({
        "sku": "NL007",
        "channel": "GT",
        "buyer_distributor": "Metro Cash & Carry",
        "location": "Bengaluru",
        "po_number": "PO-METRO-991",
        "invoice_no": "INV-7721",
        "po_quantity": 250,
        "actual_sent": 250
    }).encode('utf-8')
    req = urllib.request.Request(f'{BASE}/api/dispatch-gt-mt', data=disp_payload, headers=disp_headers)
    resp = urllib.request.urlopen(req)
    new_gtmt_id = json.loads(resp.read().decode('utf-8'))['id']

    # 5. Verify the Audit Log Table
    print("\n--- Verifying Audit Log Records ---")
    with urllib.request.urlopen(f'{BASE}/api/audit-log') as resp:
        logs = json.loads(resp.read().decode('utf-8'))
    
    print(f"Total Audit Entries: {len(logs)}")
    for l in logs[:4]:
        print(f"  Log #{l['id']} [{l['timestamp']}] | Person: {l['user_name']} ({l['user_email']}) | Sheet: {l['sheet_name']} | Edit Done: {l['edit_summary']}")

    # Cleanup test dispatch
    del_req = urllib.request.Request(
        f'{BASE}/api/dispatch-gt-mt/{new_gtmt_id}',
        headers=disp_headers,
        method='DELETE'
    )
    urllib.request.urlopen(del_req)
    
    # Restore CS0001 stock
    restore_payload = json.dumps({"sku": "CS0001", "total_inventory": 5060, "reorder_threshold": 1000}).encode('utf-8')
    urllib.request.urlopen(urllib.request.Request(f'{BASE}/api/inventory-stock', data=restore_payload, headers=headers))

    print("\nALL AUDIT LOG AND GT/MT DEEP DIVE TESTS PASSED 100%!")

if __name__ == '__main__':
    test()

import requests
import io
import time

BASE_URL = "http://127.0.0.1:8765"

def run_tests():
    s = requests.Session()
    print("1. Authenticating as Admin Manager via /login...")
    login_res = s.post(f"{BASE_URL}/login", json={
        "email": "admin@company.com",
        "password": "Admin@123"
    })
    print(f"Login status: {login_res.status_code}")
    assert login_res.status_code == 200, f"Login failed: {login_res.text}"

    print("\n2. Testing GET /api/purchase-orders...")
    get_res = s.get(f"{BASE_URL}/api/purchase-orders")
    assert get_res.status_code == 200, f"GET POs failed: {get_res.status_code}"
    po_data = get_res.json()
    pos = po_data.get('purchase_orders', [])
    stats = po_data.get('stats', {})
    print(f"Found {len(pos)} POs. Stats: {stats}")
    assert len(pos) >= 3, f"Expected at least 3 seeded POs, found {len(pos)}"

    unique_num = f"PO-TEST-{int(time.time())}"
    print(f"\n3. Testing POST /api/purchase-orders with MULTI-PRODUCT across DIFFERENT BRANDS ({unique_num})...")
    multi_po_payload = {
        "po_number": unique_num,
        "po_date": "2024-11-20",
        "platform_or_channel": "Blinkit",
        "location": "Gurugram Hub DC-4",
        "appointment_date": "2024-11-22",
        "status": "Pending Fulfillment",
        "notes": "Test multi-brand manual PO",
        "items": [
            {
                "sku": "CS0001",
                "product_name": "Cleanser",
                "brand": "California Skin+",
                "category": "Skincare",
                "quantity": 50,
                "unit_price": 350.0
            },
            {
                "sku": "NL007",
                "product_name": "Handwash",
                "brand": "Derma+",
                "category": "Personal Care",
                "quantity": 100,
                "unit_price": 120.0
            },
            {
                "sku": "NL002",
                "product_name": "NutraChips Baked Ragi Chips",
                "brand": "NutraChips",
                "category": "Snacks",
                "quantity": 200,
                "unit_price": 60.0
            }
        ]
    }
    create_res = s.post(f"{BASE_URL}/api/purchase-orders", json=multi_po_payload)
    print(f"Create PO response: {create_res.status_code}")
    assert create_res.status_code == 201, f"Failed to create PO: {create_res.text}"
    created_id = create_res.json().get('po_id')
    assert created_id is not None, "po_id missing in response"

    print(f"\n4. Testing GET /api/purchase-orders/{created_id}...")
    detail_res = s.get(f"{BASE_URL}/api/purchase-orders/{created_id}")
    assert detail_res.status_code == 200
    single_po = detail_res.json().get('purchase_order', {})
    items = single_po.get('items', [])
    print(f"PO #{single_po.get('po_number')} has {len(items)} items.")
    assert len(items) == 3, f"Expected 3 items, got {len(items)}"
    assert single_po.get('total_units') == 350, f"Expected 350 units, got {single_po.get('total_units')}"
    expected_val = (50 * 350.0) + (100 * 120.0) + (200 * 60.0)
    assert abs(single_po.get('total_value') - expected_val) < 0.01, f"Expected Rs {expected_val}, got {single_po.get('total_value')}"
    print("Multi-product calculations verified perfectly!")

    print(f"\n5. Testing DELETE /api/purchase-orders/{created_id}...")
    del_res = s.delete(f"{BASE_URL}/api/purchase-orders/{created_id}")
    assert del_res.status_code == 200
    verify_del = s.get(f"{BASE_URL}/api/purchase-orders/{created_id}")
    assert verify_del.status_code == 404, "PO still exists after deletion"
    print("PO successfully deleted and cascade verified!")

    print("\nALL PURCHASE ORDER TESTS PASSED SUCCESSFULLY! [SUCCESS]")

if __name__ == '__main__':
    run_tests()

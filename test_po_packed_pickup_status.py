import unittest
import json
import sqlite3
from app import app, get_db

class TestPOPackedPickupStatus(unittest.TestCase):
    def setUp(self):
        self.app = app.test_client()
        self.app.testing = True

    def login_as(self, email):
        with self.app.session_transaction() as sess:
            sess['user_email'] = email
            sess['user_name'] = 'Test User'

    def test_po_logistics_workflow(self):
        # 1. Login as admin
        self.login_as('admin@company.com')

        import time
        test_po_num = f'PO-TEST-LOG-{int(time.time())}'
        payload = {
            'po_number': test_po_num,
            'po_date': '2026-09-23',
            'platform_or_channel': 'Zepto',
            'location': 'Mumbai Hub',
            'appointment_date': '2026-09-25',
            'is_packed': 0,
            'is_pickup_ready': 0,
            'items': [
                {
                    'sku': 'SKU-LOG-1',
                    'product_name': 'Logistics Test Item A',
                    'brand': 'CompanyBrand',
                    'quantity': 50,
                    'unit_price': 100.0
                }
            ]
        }
        res = self.app.post('/api/purchase-orders', data=json.dumps(payload), content_type='application/json')
        self.assertEqual(res.status_code, 201)
        res_data = res.get_json()
        self.assertEqual(res_data.get('status'), 'success')
        po_id = res_data.get('po_id')
        self.assertIsNotNone(po_id)

        # 3. Retrieve PO details and verify initial flags
        res = self.app.get(f'/api/purchase-orders/{po_id}')
        self.assertEqual(res.status_code, 200)
        po_detail = res.get_json().get('purchase_order')
        self.assertEqual(po_detail['is_packed'], 0)
        self.assertEqual(po_detail['packed_yes_no'], 'No')
        self.assertEqual(po_detail['is_pickup_ready'], 0)
        self.assertEqual(po_detail['pickup_ready_yes_no'], 'No')
        self.assertEqual(po_detail['is_dispatched'], 0)
        self.assertEqual(po_detail['dispatch_status_yes_no'], 'No')

        # 4. Toggle packed to 1
        res = self.app.post(f'/api/purchase-orders/{po_id}/toggle', data=json.dumps({
            'field': 'packed',
            'value': 1
        }), content_type='application/json')
        self.assertEqual(res.status_code, 200)
        t_data = res.get_json()
        self.assertEqual(t_data['is_packed'], 1)
        self.assertEqual(t_data['packed_yes_no'], 'Yes')

        # 5. Toggle pickup_ready to 1 (Packed must remain 1)
        res = self.app.post(f'/api/purchase-orders/{po_id}/toggle', data=json.dumps({
            'field': 'pickup_ready',
            'value': 1
        }), content_type='application/json')
        self.assertEqual(res.status_code, 200)
        t_data = res.get_json()
        self.assertEqual(t_data['is_pickup_ready'], 1)
        self.assertEqual(t_data['pickup_ready_yes_no'], 'Yes')
        self.assertEqual(t_data['is_packed'], 1)

        # 6. Toggle packed to 0 -> should auto-reset pickup_ready to 0
        res = self.app.post(f'/api/purchase-orders/{po_id}/toggle', data=json.dumps({
            'field': 'packed',
            'value': 0
        }), content_type='application/json')
        self.assertEqual(res.status_code, 200)
        t_data = res.get_json()
        self.assertEqual(t_data['is_packed'], 0)
        self.assertEqual(t_data['packed_yes_no'], 'No')
        self.assertEqual(t_data['is_pickup_ready'], 0)
        self.assertEqual(t_data['pickup_ready_yes_no'], 'No')

        # 7. Setting pickup_ready to 1 directly should cascade packed to 1
        res = self.app.post(f'/api/purchase-orders/{po_id}/toggle', data=json.dumps({
            'field': 'pickup_ready',
            'value': 1
        }), content_type='application/json')
        self.assertEqual(res.status_code, 200)
        t_data = res.get_json()
        self.assertEqual(t_data['is_packed'], 1)
        self.assertEqual(t_data['is_pickup_ready'], 1)

        # 8. Test list endpoint and filters
        res = self.app.get('/api/purchase-orders?packed=Yes')
        self.assertEqual(res.status_code, 200)
        orders = res.get_json().get('purchase_orders', [])
        self.assertTrue(any(o['id'] == po_id for o in orders))
        self.assertTrue(all(o['is_packed'] == 1 for o in orders))

        res = self.app.get('/api/purchase-orders?packed=No')
        self.assertEqual(res.status_code, 200)
        orders_unpacked = res.get_json().get('purchase_orders', [])
        self.assertFalse(any(o['id'] == po_id for o in orders_unpacked))

        res = self.app.get('/api/purchase-orders?dispatch_status=No')
        self.assertEqual(res.status_code, 200)
        orders_undispatched = res.get_json().get('purchase_orders', [])
        self.assertTrue(any(o['id'] == po_id for o in orders_undispatched))

        # 9. Toggle dispatch to 1
        res = self.app.post(f'/api/purchase-orders/{po_id}/toggle', data=json.dumps({
            'field': 'dispatched',
            'value': 1
        }), content_type='application/json')
        self.assertEqual(res.status_code, 200)
        t_data = res.get_json()
        self.assertEqual(t_data['is_dispatched'], 1)
        self.assertEqual(t_data['dispatch_status_yes_no'], 'Yes')

        # 10. Clean up test PO
        res = self.app.delete(f'/api/purchase-orders/{po_id}')
        self.assertEqual(res.status_code, 200)
        print('All PO Packed, Pickup (Packed & Ready), and Dispatch Status tests passed successfully!')

if __name__ == '__main__':
    unittest.main()

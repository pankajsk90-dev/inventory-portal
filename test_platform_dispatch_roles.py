import unittest
import json
import sqlite3
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from app import app, get_db, init_schema

class TestPlatformDispatchRoles(unittest.TestCase):
    def setUp(self):
        self.client = app.test_client()
        init_schema()

        db = get_db()
        # Clean test records
        db.execute("DELETE FROM purchase_orders WHERE po_number LIKE 'PO-TEST-%'")
        db.execute("DELETE FROM po_items WHERE po_number LIKE 'PO-TEST-%'")
        db.execute("DELETE FROM dispatch_online WHERE po_number LIKE 'PO-TEST-%'")
        db.execute("DELETE FROM dispatch_gt_mt WHERE po_number LIKE 'PO-TEST-%'")

        # Dispatch Lead in permissions
        db.execute("""
            INSERT OR REPLACE INTO permissions (role_key, authorized_email, description)
            VALUES ('DISPATCH_MANAGER', 'dispatch_mgr@company.com', 'Dispatch Lead')
        """)
        # Platform Managers in platform_assignments
        db.execute("""
            INSERT OR REPLACE INTO platform_assignments (platform_key, channel_type, assigned_name, assigned_email)
            VALUES ('Blinkit', 'online', 'Rahul', 'rahul@company.com')
        """)
        db.execute("""
            INSERT OR REPLACE INTO platform_assignments (platform_key, channel_type, assigned_name, assigned_email)
            VALUES ('Zepto', 'online', 'Priya', 'priya@company.com')
        """)
        db.commit()

    def login_as(self, email, name="Test User"):
        with self.client.session_transaction() as sess:
            sess['user_email'] = email
            sess['user_name'] = name

    def test_po_creation_permission_checks(self):
        po_blinkit = {
            'po_number': 'PO-TEST-BLK-001',
            'po_date': '2026-09-23',
            'platform_or_channel': 'Blinkit',
            'location': 'Gurgaon FC',
            'appointment_date': '2026-09-25',
            'status': 'Pending Fulfillment',
            'items': [
                { 'sku': 'SKU-BLK-1', 'product_name': 'Test Snack 100g', 'brand': 'BrandA', 'quantity': 50, 'unit_price': 40.0 }
            ]
        }

        # 1. Rahul (Blinkit) creates PO for Blinkit -> Should SUCCEED
        self.login_as('rahul@company.com', 'Rahul (Blinkit Lead)')
        res = self.client.post('/api/purchase-orders', json=po_blinkit)
        self.assertEqual(res.status_code, 201, f'Rahul creating Blinkit PO failed: {res.data}')
        data = json.loads(res.data)
        po_id = data['po_id']

        # 2. Rahul (Blinkit) tries to create PO for Zepto -> Should be FORBIDDEN (403)
        po_zepto = {
            'po_number': 'PO-TEST-ZPT-001',
            'po_date': '2026-09-23',
            'platform_or_channel': 'Zepto',
            'items': [
                { 'sku': 'SKU-ZPT-1', 'product_name': 'Test Juice 200ml', 'brand': 'BrandB', 'quantity': 30, 'unit_price': 50.0 }
            ]
        }
        res_fail = self.client.post('/api/purchase-orders', json=po_zepto)
        self.assertEqual(res_fail.status_code, 403, 'Rahul should NOT be allowed to create Zepto PO')

        # 3. Dispatch Manager tries to create a PO -> Should be FORBIDDEN (403)
        self.login_as('dispatch_mgr@company.com', 'Dispatch Lead')
        res_disp = self.client.post('/api/purchase-orders', json=po_blinkit)
        self.assertEqual(res_disp.status_code, 403, 'Dispatch Manager should NOT be allowed to create POs')

        # Clean up
        self.login_as('admin@company.com', 'Admin')
        self.client.delete(f'/api/purchase-orders/{po_id}')

    def test_dispatch_manager_courier_booking_and_grn_restrictions(self):
        # 1. Platform Manager creates PO
        po_payload = {
            'po_number': 'PO-TEST-BLK-002',
            'po_date': '2026-09-23',
            'platform_or_channel': 'Blinkit',
            'location': 'Noida Hub',
            'items': [
                { 'sku': 'SKU-BLK-A', 'product_name': 'Product A', 'brand': 'BrandA', 'quantity': 100, 'unit_price': 25.0 },
                { 'sku': 'SKU-BLK-B', 'product_name': 'Product B', 'brand': 'BrandB', 'quantity': 50, 'unit_price': 60.0 }
            ]
        }
        self.login_as('rahul@company.com', 'Rahul (Blinkit Lead)')
        create_res = self.client.post('/api/purchase-orders', json=po_payload)
        self.assertEqual(create_res.status_code, 201)
        po_id = json.loads(create_res.data)['po_id']

        # 2. Dispatch Manager fulfills / books courier for this PO
        fulfill_payload = {
            'po_number': 'PO-TEST-BLK-002',
            'dispatch_type': 'online',
            'courier_name': 'BlueDart Express',
            'tracking_id': 'BD-88997711',
            'courier_charges': 450.0,
            'items': [
                { 'sku': 'SKU-BLK-A', 'actual_sent': 100 },
                { 'sku': 'SKU-BLK-B', 'actual_sent': 50 }
            ]
        }
        self.login_as('dispatch_mgr@company.com', 'Dispatch Lead')
        fulfill_res = self.client.post('/api/dispatch/fulfill-po', json=fulfill_payload)
        self.assertEqual(fulfill_res.status_code, 200, f'Dispatch manager fulfillment failed: {fulfill_res.data}')
        fulfill_data = json.loads(fulfill_res.data)
        self.assertEqual(fulfill_data['status'], 'success')

        po_check = self.client.get(f'/api/purchase-orders/{po_id}')
        self.assertEqual(po_check.get_json()['purchase_order']['status'], 'Fully Dispatched')

        # 3. Dispatch Manager tries to mark GRN Done -> Should be FORBIDDEN (403)
        grn_payload = {
            'po_id': po_id,
            'grn_number': 'GRN-BLK-9988',
            'grn_date': '2026-09-24',
            'grn_notes': 'Attempt by dispatch manager'
        }
        res_grn_disp = self.client.post(f'/api/purchase-orders/{po_id}/grn', json=grn_payload)
        self.assertEqual(res_grn_disp.status_code, 403, 'Dispatch manager must NOT be allowed to mark GRN Done')

        # 4. Another Platform Manager (Priya - Zepto) tries to mark GRN Done for Blinkit PO -> FORBIDDEN (403)
        self.login_as('priya@company.com', 'Priya (Zepto Lead)')
        res_grn_priya = self.client.post(f'/api/purchase-orders/{po_id}/grn', json=grn_payload)
        self.assertEqual(res_grn_priya.status_code, 403, 'Priya must NOT be allowed to mark GRN Done for Blinkit PO')

        # 5. Assigned Platform Manager (Rahul - Blinkit) marks GRN Done -> Should SUCCEED (200)
        self.login_as('rahul@company.com', 'Rahul (Blinkit Lead)')
        res_grn_rahul = self.client.post(f'/api/purchase-orders/{po_id}/grn', json=grn_payload)
        self.assertEqual(res_grn_rahul.status_code, 200, f'Rahul marking GRN failed: {res_grn_rahul.data}')
        rahul_data = json.loads(res_grn_rahul.data)
        self.assertEqual(rahul_data['status'], 'success')
        self.assertEqual(rahul_data['grn_number'], 'GRN-BLK-9988')

        # Verify updated PO details via GET
        po_after_grn = self.client.get(f'/api/purchase-orders/{po_id}').get_json()['purchase_order']
        self.assertEqual(po_after_grn['status'], 'GRN Done')
        self.assertEqual(po_after_grn['grn_number'], 'GRN-BLK-9988')
        self.assertEqual(po_after_grn['grn_done_by'], 'rahul@company.com')

        # Verify linked online dispatch records are updated to Delivered
        db = get_db()
        items = db.execute("SELECT dispatch_status FROM dispatch_online WHERE po_number = 'PO-TEST-BLK-002'").fetchall()
        for it in items:
            self.assertEqual(it['dispatch_status'], 'Delivered')

        # Clean up
        self.login_as('admin@company.com', 'Admin')
        self.client.delete(f'/api/purchase-orders/{po_id}')

if __name__ == '__main__':
    unittest.main()

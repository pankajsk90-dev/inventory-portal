import unittest
import json
import sqlite3
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from app import app, get_db

class TestPODispatchLink(unittest.TestCase):
    def setUp(self):
        self.client = app.test_client()
        with self.client.session_transaction() as sess:
            sess['user_email'] = 'admin@company.com'
            sess['user_name'] = 'Admin Manager'
        conn = get_db()
        cur = conn.cursor()
        cur.execute("DELETE FROM purchase_orders WHERE po_number LIKE 'PO-TEST-%'")
        cur.execute("DELETE FROM po_items WHERE po_number LIKE 'PO-TEST-%'")
        cur.execute("DELETE FROM dispatch_online WHERE po_number LIKE 'PO-TEST-%'")
        cur.execute("DELETE FROM dispatch_gt_mt WHERE po_number LIKE 'PO-TEST-%'")
        conn.commit()
        conn.close()

    def test_po_auto_routes_to_online_dispatch_and_syncs(self):
        po_payload = {
            'po_number': 'PO-TEST-BLK-999',
            'po_date': '2026-09-23',
            'platform_or_channel': 'Blinkit',
            'location': 'Blinkit Hub Delhi',
            'appointment_date': '2026-09-25',
            'status': 'Pending Fulfillment',
            'notes': 'Test online PO linking',
            'items': [
                {
                    'sku': 'CS0001',
                    'product_name': 'Cleanser',
                    'brand': 'California Skin+',
                    'category': 'Skincare',
                    'ordered_quantity': 50,
                    'unit_price': 220.0
                },
                {
                    'sku': 'NL001',
                    'product_name': 'NutraChips Rock Salt Banana Chips',
                    'brand': 'NutraChips',
                    'category': 'Chips',
                    'ordered_quantity': 40,
                    'unit_price': 65.0
                }
            ]
        }
        res = self.client.post('/api/purchase-orders', json=po_payload)
        self.assertEqual(res.status_code, 201)
        data = res.get_json()
        po_id = data['po_id']

        conn = get_db()
        cur = conn.cursor()
        rows = cur.execute('SELECT * FROM dispatch_online WHERE po_number = ?', ('PO-TEST-BLK-999',)).fetchall()
        self.assertEqual(len(rows), 2)
        for r in rows:
            self.assertEqual(r['actual_sent'], 0)
            self.assertEqual(r['dispatch_status'], 'Pending Booking')
            self.assertEqual(r['platform'], 'Blinkit')
        
        item1 = rows[0]
        courier_payload = {
            'actual_sent': item1['po_quantity'],
            'courier_name': 'Delhivery',
            'tracking_id': 'DLV-123456',
            'courier_charges': 250.0,
            'dispatch_status': 'Dispatched',
            'eway_bill_no': 'EWAY-9988',
            'shipping_notes': 'Dispatched from Bay 3'
        }
        res_courier = self.client.post(f'/api/dispatch-online/{item1["id"]}/courier', json=courier_payload)
        self.assertEqual(res_courier.status_code, 200)

        po_row = cur.execute('SELECT status FROM purchase_orders WHERE id = ?', (po_id,)).fetchone()
        self.assertEqual(po_row['status'], 'Partially Dispatched')

        item2 = rows[1]
        courier_payload2 = {
            'actual_sent': item2['po_quantity'],
            'courier_name': 'Shadowfax',
            'tracking_id': 'SFX-789012',
            'courier_charges': 200.0,
            'dispatch_status': 'Dispatched'
        }
        res_courier2 = self.client.post(f'/api/dispatch-online/{item2["id"]}/courier', json=courier_payload2)
        self.assertEqual(res_courier2.status_code, 200)

        po_row = cur.execute('SELECT status FROM purchase_orders WHERE id = ?', (po_id,)).fetchone()
        self.assertEqual(po_row['status'], 'Fully Dispatched')

        cur.execute('DELETE FROM dispatch_online WHERE po_number = ?', ('PO-TEST-BLK-999',))
        cur.execute('DELETE FROM po_items WHERE po_id = ?', (po_id,))
        cur.execute('DELETE FROM purchase_orders WHERE id = ?', (po_id,))
        conn.commit()
        conn.close()

    def test_po_auto_routes_to_gtmt_dispatch_and_cleanup(self):
        po_payload = {
            'po_number': 'PO-TEST-GT-888',
            'po_date': '2026-09-23',
            'platform_or_channel': 'GT',
            'location': 'Pune Depot',
            'distributor_name': 'Kothari Traders',
            'appointment_date': '2026-09-26',
            'status': 'Pending Fulfillment',
            'items': [
                {
                    'sku': 'CS0001',
                    'product_name': 'Cleanser',
                    'brand': 'California Skin+',
                    'category': 'Skincare',
                    'ordered_quantity': 100,
                    'unit_price': 220.0
                }
            ]
        }
        res = self.client.post('/api/purchase-orders', json=po_payload)
        self.assertEqual(res.status_code, 201)
        data = res.get_json()
        po_id = data['po_id']

        conn = get_db()
        cur = conn.cursor()
        rows = cur.execute('SELECT * FROM dispatch_gt_mt WHERE po_number = ?', ('PO-TEST-GT-888',)).fetchall()
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]['channel'], 'GT')
        self.assertEqual(rows[0]['buyer_distributor'], 'Kothari Traders')
        self.assertEqual(rows[0]['actual_sent'], 0)
        self.assertEqual(rows[0]['dispatch_status'], 'Pending Booking')
        conn.close()

        del_res = self.client.delete(f'/api/purchase-orders/{po_id}')
        self.assertEqual(del_res.status_code, 200)

        conn = get_db()
        cur = conn.cursor()
        rem_rows = cur.execute('SELECT * FROM dispatch_gt_mt WHERE po_number = ?', ('PO-TEST-GT-888',)).fetchall()
        self.assertEqual(len(rem_rows), 0)
        conn.close()

if __name__ == '__main__':
    unittest.main()

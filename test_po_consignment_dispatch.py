import unittest
import io
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from app import app, get_db

class TestPOConsignmentDispatch(unittest.TestCase):
    def setUp(self):
        self.client = app.test_client()
        with self.client.session_transaction() as sess:
            sess['user_email'] = 'admin@company.com'
            sess['user_name'] = 'Admin Manager'
        conn = get_db()
        cur = conn.cursor()
        cur.execute("DELETE FROM purchase_orders WHERE po_number LIKE 'PO-CSG-%'")
        cur.execute("DELETE FROM po_items WHERE po_number LIKE 'PO-CSG-%'")
        cur.execute("DELETE FROM dispatch_online WHERE po_number LIKE 'PO-CSG-%'")
        cur.execute("DELETE FROM dispatch_gt_mt WHERE po_number LIKE 'PO-CSG-%'")
        conn.commit()
        conn.close()

    def test_online_po_consignment_fulfillment_and_single_doc_upload(self):
        # 1. Create multi-item PO for Blinkit
        po_payload = {
            'po_number': 'PO-CSG-ONLINE-101',
            'po_date': '2026-09-23',
            'platform_or_channel': 'Blinkit',
            'location': 'Blinkit Hub Mumbai',
            'appointment_date': '2026-09-26',
            'status': 'Pending Fulfillment',
            'items': [
                {
                    'sku': 'CS0001',
                    'product_name': 'Cleanser',
                    'brand': 'California Skin+',
                    'category': 'Skincare',
                    'ordered_quantity': 60,
                    'unit_price': 220.0
                },
                {
                    'sku': 'NL001',
                    'product_name': 'NutraChips Banana',
                    'brand': 'NutraChips',
                    'category': 'Chips',
                    'ordered_quantity': 40,
                    'unit_price': 65.0
                }
            ]
        }
        res = self.client.post('/api/purchase-orders', json=po_payload)
        self.assertEqual(res.status_code, 201)

        # Verify 2 rows in dispatch_online with actual_sent = 0
        conn = get_db()
        cur = conn.cursor()
        rows = cur.execute('SELECT * FROM dispatch_online WHERE po_number = ?', ('PO-CSG-ONLINE-101',)).fetchall()
        self.assertEqual(len(rows), 2)
        for r in rows:
            self.assertEqual(r['actual_sent'], 0)
            self.assertEqual(r['dispatch_status'], 'Pending Booking')
        conn.close()

        # 2. Attach PO Copy for the whole consignment using single upload endpoint
        file_bytes = io.BytesIO(b'%PDF-1.4 Mock PO PDF Content')
        doc_res = self.client.post('/api/dispatch/upload-docs', data={
            'file': (file_bytes, 'Blinkit_Consignment_PO.pdf'),
            'dispatch_type': 'online',
            'doc_type': 'po',
            'po_number': 'PO-CSG-ONLINE-101'
        }, content_type='multipart/form-data')
        self.assertEqual(doc_res.status_code, 200)

        # Verify BOTH items in dispatch_online now have po_doc_name and po_doc_url!
        conn = get_db()
        cur = conn.cursor()
        rows = cur.execute('SELECT * FROM dispatch_online WHERE po_number = ?', ('PO-CSG-ONLINE-101',)).fetchall()
        self.assertEqual(len(rows), 2)
        for r in rows:
            self.assertIsNotNone(r['po_doc_url'])
            self.assertEqual(r['po_doc_name'], 'Blinkit_Consignment_PO.pdf')
        conn.close()

        # 3. Fulfill the entire consignment at once
        fulfill_payload = {
            'dispatch_type': 'online',
            'po_number': 'PO-CSG-ONLINE-101',
            'courier_name': 'Delhivery Express',
            'tracking_id': 'DEL77889911',
            'courier_charges': 1500.0,
            'dispatch_status': 'Dispatched',
            'eway_bill_no': 'EWAY12345678',
            'shipping_notes': '2 Boxes, Inspected',
            'items': [
                {'sku': 'CS0001', 'actual_sent': 60},
                {'sku': 'NL001', 'actual_sent': 40}
            ]
        }
        f_res = self.client.post('/api/dispatch/fulfill-po', json=fulfill_payload)
        self.assertEqual(f_res.status_code, 200)

        # Verify all dispatch rows and po_items got updated
        conn = get_db()
        cur = conn.cursor()
        rows = cur.execute('SELECT * FROM dispatch_online WHERE po_number = ? ORDER BY sku', ('PO-CSG-ONLINE-101',)).fetchall()
        self.assertEqual(len(rows), 2)
        total_freight = 0
        for r in rows:
            self.assertEqual(r['courier_name'], 'Delhivery Express')
            self.assertEqual(r['tracking_id'], 'DEL77889911')
            self.assertEqual(r['dispatch_status'], 'Dispatched')
            self.assertEqual(r['eway_bill_no'], 'EWAY12345678')
            total_freight += r['courier_charges']
            if r['sku'] == 'CS0001':
                self.assertEqual(r['actual_sent'], 60)
            elif r['sku'] == 'NL001':
                self.assertEqual(r['actual_sent'], 40)
        self.assertAlmostEqual(total_freight, 1500.0, places=1)

        # Parent PO status should now be 'Fully Dispatched'
        parent_po = cur.execute('SELECT * FROM purchase_orders WHERE po_number = ?', ('PO-CSG-ONLINE-101',)).fetchone()
        self.assertEqual(parent_po['status'], 'Fully Dispatched')
        conn.close()

    def test_gt_po_consignment_fulfillment_and_invoice_upload(self):
        # 1. Create GT PO
        po_payload = {
            'po_number': 'PO-CSG-GT-202',
            'po_date': '2026-09-23',
            'platform_or_channel': 'GT',
            'buyer_distributor': 'Apex Distributors Pune',
            'location': 'Pune Hub',
            'status': 'Pending Fulfillment',
            'items': [
                {
                    'sku': 'CS0001',
                    'product_name': 'Cleanser',
                    'brand': 'California Skin+',
                    'category': 'Skincare',
                    'ordered_quantity': 30,
                    'unit_price': 220.0
                }
            ]
        }
        res = self.client.post('/api/purchase-orders', json=po_payload)
        self.assertEqual(res.status_code, 201)

        # 2. Attach Invoice Copy
        file_bytes = io.BytesIO(b'%PDF-1.4 Mock Invoice Content')
        doc_res = self.client.post('/api/dispatch/upload-docs', data={
            'file': (file_bytes, 'Apex_Invoice_INV202.pdf'),
            'dispatch_type': 'gtmt',
            'doc_type': 'invoice',
            'po_number': 'PO-CSG-GT-202'
        }, content_type='multipart/form-data')
        self.assertEqual(doc_res.status_code, 200)

        # 3. Fulfill GT consignment
        fulfill_payload = {
            'dispatch_type': 'gtmt',
            'po_number': 'PO-CSG-GT-202',
            'courier_name': 'Gati KWE',
            'tracking_id': 'LR-PUNE-8889',
            'courier_charges': 850.0,
            'dispatch_status': 'In Transit',
            'eway_bill_no': 'EWAY-GT-999',
            'shipping_notes': 'Delivered via Pune Express',
            'items': [
                {'sku': 'CS0001', 'actual_sent': 30}
            ]
        }
        f_res = self.client.post('/api/dispatch/fulfill-po', json=fulfill_payload)
        self.assertEqual(f_res.status_code, 200)

        conn = get_db()
        cur = conn.cursor()
        r = cur.execute('SELECT * FROM dispatch_gt_mt WHERE po_number = ?', ('PO-CSG-GT-202',)).fetchone()
        self.assertEqual(r['courier_name'], 'Gati KWE')
        self.assertEqual(r['tracking_id'], 'LR-PUNE-8889')
        self.assertEqual(r['actual_sent'], 30)
        self.assertEqual(r['invoice_doc_name'], 'Apex_Invoice_INV202.pdf')
        conn.close()

if __name__ == '__main__':
    unittest.main()

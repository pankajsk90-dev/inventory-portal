import unittest
import os
import io
import json
from app import app, get_db

class TestBlinkitPredictiveEngine(unittest.TestCase):
    def setUp(self):
        app.config['TESTING'] = True
        app.config['WTF_CSRF_ENABLED'] = False
        self.client = app.test_client()
        # Login as Admin
        with self.client as c:
            c.post('/login', data={'email': 'admin@company.com', 'password': 'Admin@123'})

    def test_01_unauthenticated_api_returns_401(self):
        anon_client = app.test_client()
        res = anon_client.get('/api/blinkit/replenishment-model')
        self.assertEqual(res.status_code, 401)
        res_settings = anon_client.get('/api/blinkit/settings')
        self.assertEqual(res_settings.status_code, 401)

    def test_02_replenishment_model_kpis_and_rules(self):
        with self.client as c:
            res = c.get('/api/blinkit/replenishment-model')
            self.assertEqual(res.status_code, 200)
            data = res.get_json()
            self.assertEqual(data['status'], 'success')
            kpis = data['kpis']
            
            # Verify data presence from auto-seeding
            self.assertGreater(kpis['total_facilities'], 0)
            self.assertGreater(kpis['total_items_monitored'], 0)
            self.assertGreater(kpis['total_sales_orders'], 0)
            self.assertIn('critical_count', kpis)
            self.assertIn('overstocked_count', kpis)
            self.assertIn('excess_capital_at_risk', kpis)

            # Check mathematical rules across rows
            rows = data['rows']
            self.assertGreater(len(rows), 0)
            
            for r in rows:
                lead_time = r['lead_time_days']
                safety_stock = r['safety_stock_days']
                target_max = r['target_max_days']
                pipeline_doi = r['pipeline_doi']
                drr = r['drr']
                status = r['status']
                rec_po = r['recommended_po_qty']

                # Delivery window verification: Lead time within 7-10 days window, Safety stock 3 days
                self.assertIn(lead_time, [7, 8, 9, 10])
                self.assertEqual(safety_stock, 3)
                self.assertEqual(target_max, 21)

                if drr > 0:
                    if pipeline_doi < lead_time:
                        self.assertEqual(status, 'CRITICAL')
                        self.assertGreater(rec_po, 0)
                    elif pipeline_doi <= (lead_time + safety_stock):
                        self.assertEqual(status, 'REORDER_DUE')
                        self.assertGreater(rec_po, 0)
                    elif pipeline_doi <= target_max:
                        self.assertEqual(status, 'HEALTHY')
                        self.assertEqual(rec_po, 0)
                    else:
                        self.assertEqual(status, 'OVERSTOCKED')
                        self.assertEqual(rec_po, 0)

    def test_03_zero_overlap_sales_deduplication(self):
        with self.client as c:
            # Prepare a sample CSV with 2 orders, one existing and one brand new
            csv_content = """Order Id,Order Date,Item Id,Product Name,Brand Name,UPC,Supply City,Supply State,Customer City,Customer State,Quantity,MRP (Rs),Selling Price (Rs),Total Gross Bill Amount,Order Status
1003759902636,2026-09-20,1000000000001,Test Serum 30ml,Brand B,UPC999,Mumbai,MH,Mumbai,MH,2,599,499,998,DELIVERED
ORDER_TEST_UNIQUE_9999,2026-09-27,1000000000001,Test Serum 30ml,Brand B,UPC999,Mumbai,MH,Mumbai,MH,1,599,499,499,DELIVERED
"""
            # First upload
            res1 = c.post('/api/blinkit/upload-sales', data={
                'file': (io.BytesIO(csv_content.encode('utf-8')), 'test_sales_1.csv')
            }, content_type='multipart/form-data')
            self.assertEqual(res1.status_code, 200)
            d1 = res1.get_json()
            self.assertEqual(d1['total_orders_in_file'], 2)
            # Second upload of the EXACT SAME FILE
            res2 = c.post('/api/blinkit/upload-sales', data={
                'file': (io.BytesIO(csv_content.encode('utf-8')), 'test_sales_2.csv')
            }, content_type='multipart/form-data')
            self.assertEqual(res2.status_code, 200)
            d2 = res2.get_json()
            self.assertEqual(d2['total_orders_in_file'], 2)
            # Crucial verification: new orders inserted must be 0, duplicates skipped must be 2
            self.assertEqual(d2['new_orders_inserted'], 0)
            self.assertEqual(d2['existing_orders_skipped'], 2)

    def test_04_facility_settings_update(self):
        with self.client as c:
            # Update settings for a specific facility
            res = c.post('/api/blinkit/settings', json={
                'facility_name': 'Mumbai M12 - Feeder Warehouse',
                'lead_time_days': 9,
                'safety_stock_days': 4,
                'target_max_days': 20
            })
            self.assertEqual(res.status_code, 200)
            settings = res.get_json()['settings']
            m12 = next((s for s in settings if s['facility_name'] == 'Mumbai M12 - Feeder Warehouse'), None)
            self.assertIsNotNone(m12)
            self.assertEqual(m12['lead_time_days'], 9)
            self.assertEqual(m12['safety_stock_days'], 4)
            self.assertEqual(m12['target_max_days'], 20)

            # Restore default
            c.post('/api/blinkit/settings', json={
                'facility_name': 'Mumbai M12 - Feeder Warehouse',
                'lead_time_days': 8,
                'safety_stock_days': 3,
                'target_max_days': 21
            })

if __name__ == '__main__':
    unittest.main()

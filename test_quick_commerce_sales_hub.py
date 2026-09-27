import unittest
import os
import sqlite3
import datetime
import app

class TestQuickCommerceSalesHub(unittest.TestCase):
    def setUp(self):
        self.app = app.app.test_client()
        self.conn = app.get_db()
        self.cur = self.conn.cursor()
        self.brain_uploads = os.path.join(os.path.expanduser('~'), '.gemini', 'antigravity', 'brain', 'c18c0bb2-1dda-4e5b-89b7-d36de3bd9d79', '.user_uploaded')

    def tearDown(self):
        self.conn.close()

    def test_zepto_parsing_and_anti_overlap(self):
        z_path = os.path.join(self.brain_uploads, 'media_1790524533163.csv')
        self.assertTrue(os.path.exists(z_path), "Zepto sample CSV should exist")
        
        with open(z_path, 'rb') as f:
            content = f.read()
            
        # First ingestion
        res1 = app.parse_zepto_sales_data(content, 'test_zepto.csv', 'Tester', self.cur)
        self.conn.commit()
        count1 = self.cur.execute('SELECT COUNT(*) FROM zepto_sales_orders').fetchone()[0]
        self.assertGreater(count1, 0, "Zepto rows should be > 0")
        
        # Second ingestion of the same file (same dates, same SKUs, same cities)
        res2 = app.parse_zepto_sales_data(content, 'test_zepto.csv', 'Tester', self.cur)
        self.conn.commit()
        count2 = self.cur.execute('SELECT COUNT(*) FROM zepto_sales_orders').fetchone()[0]
        
        # Must be identical - NO DUPLICATION / OVERLAP!
        self.assertEqual(count1, count2, f"Re-uploading same date should not duplicate records! {count1} vs {count2}")
        
        # Check central sales_data deduplication
        sd_count = self.cur.execute("SELECT COUNT(*) FROM sales_data WHERE platform = 'Zepto' AND sale_date = '2026-09-25'").fetchone()[0]
        self.assertGreater(sd_count, 0)
        
        # Re-upload third time
        app.parse_zepto_sales_data(content, 'test_zepto.csv', 'Tester', self.cur)
        self.conn.commit()
        sd_count2 = self.cur.execute("SELECT COUNT(*) FROM sales_data WHERE platform = 'Zepto' AND sale_date = '2026-09-25'").fetchone()[0]
        self.assertEqual(sd_count, sd_count2, "Central sales_data should not duplicate on re-upload")

    def test_instamart_parsing_and_anti_overlap(self):
        i_path = os.path.join(self.brain_uploads, 'media_1790525376736.csv')
        self.assertTrue(os.path.exists(i_path), "Instamart sample CSV should exist")
        
        with open(i_path, 'rb') as f:
            content = f.read()
            
        # First ingestion
        res1 = app.parse_instamart_sales_data(content, 'test_instamart.csv', 'Tester', self.cur)
        self.conn.commit()
        count1 = self.cur.execute('SELECT COUNT(*) FROM instamart_sales_orders').fetchone()[0]
        self.assertGreater(count1, 0, "Instamart rows should be > 0")
        
        # Second ingestion of the same file
        res2 = app.parse_instamart_sales_data(content, 'test_instamart.csv', 'Tester', self.cur)
        self.conn.commit()
        count2 = self.cur.execute('SELECT COUNT(*) FROM instamart_sales_orders').fetchone()[0]
        
        # Must be identical - NO DUPLICATION / OVERLAP!
        self.assertEqual(count1, count2, f"Re-uploading same date should not duplicate Instamart records! {count1} vs {count2}")
        
        # Check central sales_data deduplication
        sd_count = self.cur.execute("SELECT COUNT(*) FROM sales_data WHERE platform = 'Instamart' AND sale_date = '2026-09-25'").fetchone()[0]
        self.assertGreater(sd_count, 0)
        
        # Re-upload third time
        app.parse_instamart_sales_data(content, 'test_instamart.csv', 'Tester', self.cur)
        self.conn.commit()
        sd_count2 = self.cur.execute("SELECT COUNT(*) FROM sales_data WHERE platform = 'Instamart' AND sale_date = '2026-09-25'").fetchone()[0]
        self.assertEqual(sd_count, sd_count2, "Central sales_data should not duplicate Instamart records on re-upload")

    def test_summary_apis(self):
        self.app.post('/login', data={'email': 'admin@company.com', 'password': 'Admin@123'})

        # Zepto summary API
        z_res = self.app.get('/api/zepto/sales-summary')
        self.assertEqual(z_res.status_code, 200)
        z_data = z_res.get_json()
        self.assertEqual(z_data['status'], 'success')
        self.assertIn('kpis', z_data)
        self.assertIn('cities', z_data)

        # Instamart summary API
        i_res = self.app.get('/api/instamart/sales-summary')
        self.assertEqual(i_res.status_code, 200)
        i_data = i_res.get_json()
        self.assertEqual(i_data['status'], 'success')
        self.assertIn('kpis', i_data)
        self.assertIn('records', i_data)

        # Blinkit sales tab summary API
        b_res = self.app.get('/api/blinkit/sales-tab-summary')
        self.assertEqual(b_res.status_code, 200)
        b_data = b_res.get_json()
        self.assertEqual(b_data['status'], 'success')

        # Hub summary API
        h_res = self.app.get('/api/sales/hub-summary')
        self.assertEqual(h_res.status_code, 200)
        h_data = h_res.get_json()
        self.assertEqual(h_data['status'], 'success')
        self.assertIn('zepto', h_data)
        self.assertIn('instamart', h_data)
        self.assertIn('blinkit', h_data)

    def test_instamart_excel_sales_report(self):
        real_excel_path = r'c:\Users\Admin\Downloads\13671_1790410348788.xlsx'
        if os.path.exists(real_excel_path):
            with open(real_excel_path, 'rb') as f:
                content = f.read()
            res = app.parse_instamart_sales_data(content, '13671_1790410348788.xlsx', 'Tester', self.cur)
            self.conn.commit()
            self.assertEqual(res['status'], 'success')
            self.assertEqual(res['total_rows_in_file'], 234)
            self.assertEqual(res['total_units_sold'], 277)
            self.assertEqual(res['total_gmv'], 50743.0)

if __name__ == '__main__':
    unittest.main()

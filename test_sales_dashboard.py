import unittest
import os
import sqlite3
import datetime
from app import app, get_db

class TestSalesDashboard(unittest.TestCase):
    def setUp(self):
        self.client = app.test_client()
        with self.client.session_transaction() as sess:
            sess['user_email'] = 'admin@company.com'
            sess['user_name'] = 'Admin Manager'
            
        self.conn = get_db()
        self.cur = self.conn.cursor()

    def tearDown(self):
        self.conn.close()

    def test_dashboard_analytics_unfiltered(self):
        """Test fetching full sales dashboard analytics without filters."""
        res = self.client.get('/api/sales/dashboard-analytics')
        self.assertEqual(res.status_code, 200)
        data = res.get_json()
        self.assertEqual(data.get('status'), 'success')
        
        # Verify KPI structure
        kpis = data.get('kpis', {})
        self.assertIn('total_records', kpis)
        self.assertIn('total_units', kpis)
        self.assertIn('total_revenue', kpis)
        self.assertIn('active_platforms', kpis)
        self.assertIn('active_locations', kpis)
        self.assertIn('active_products', kpis)
        self.assertIn('avg_unit_value', kpis)
        self.assertGreater(kpis['total_units'], 0)
        self.assertGreater(kpis['total_revenue'], 0)
        
        # Verify Platforms breakdown
        platforms = data.get('platforms', [])
        self.assertGreaterEqual(len(platforms), 4) # Blinkit, Zepto, Instamart, BigBasket
        plat_names = [p['platform'] for p in platforms]
        self.assertIn('Blinkit', plat_names)
        self.assertIn('Zepto', plat_names)
        self.assertIn('Instamart', plat_names)
        self.assertIn('BigBasket', plat_names)
        
        # Verify Top Products & Top Locations
        self.assertGreater(len(data.get('top_products', [])), 0)
        self.assertGreater(len(data.get('top_locations', [])), 0)
        
        # Verify Filter Options
        f_opts = data.get('filter_options', {})
        self.assertIn('platforms', f_opts)
        self.assertIn('products', f_opts)
        self.assertIn('locations', f_opts)
        self.assertIn('min_date', f_opts)
        self.assertIn('max_date', f_opts)

    def test_dashboard_platform_filter(self):
        """Test platform filtering on Zepto and BigBasket."""
        # Zepto filter
        res_z = self.client.get('/api/sales/dashboard-analytics?platform=Zepto')
        self.assertEqual(res_z.status_code, 200)
        data_z = res_z.get_json()
        for r in data_z['records']:
            self.assertEqual(r['platform'], 'Zepto')
            
        # BigBasket filter
        res_bb = self.client.get('/api/sales/dashboard-analytics?platform=BigBasket')
        self.assertEqual(res_bb.status_code, 200)
        data_bb = res_bb.get_json()
        for r in data_bb['records']:
            self.assertEqual(r['platform'], 'BigBasket')

    def test_dashboard_location_filter(self):
        """Test filtering by store_location / city."""
        res = self.client.get('/api/sales/dashboard-analytics?location=Mumbai')
        self.assertEqual(res.status_code, 200)
        data = res.get_json()
        self.assertGreater(data['kpis']['total_units'], 0)
        for r in data['records']:
            self.assertEqual(r['store_location'], 'Mumbai')

    def test_dashboard_product_filter(self):
        """Test filtering by product SKU."""
        res = self.client.get('/api/sales/dashboard-analytics?sku=CS0003')
        self.assertEqual(res.status_code, 200)
        data = res.get_json()
        self.assertGreater(data['kpis']['total_units'], 0)
        for r in data['records']:
            self.assertEqual(r['sku'], 'CS0003')

    def test_dashboard_date_range_filter(self):
        """Test filtering by start_date and end_date."""
        res = self.client.get('/api/sales/dashboard-analytics?start_date=2026-09-01&end_date=2026-09-28')
        self.assertEqual(res.status_code, 200)
        data = res.get_json()
        for r in data['records']:
            self.assertTrue(r['sale_date'] >= '2026-09-01')
            self.assertTrue(r['sale_date'] <= '2026-09-28')

    def test_dashboard_combined_multi_filter(self):
        """Test multi-dimensional filtering across Platform + Location."""
        res = self.client.get('/api/sales/dashboard-analytics?platform=Zepto&location=Mumbai')
        self.assertEqual(res.status_code, 200)
        data = res.get_json()
        for r in data['records']:
            self.assertEqual(r['platform'], 'Zepto')
            self.assertEqual(r['store_location'], 'Mumbai')

    def test_export_filtered_csv(self):
        """Test the dedicated CSV export endpoint with filters."""
        res = self.client.get('/api/sales/export-csv?platform=Zepto&location=Mumbai')
        self.assertEqual(res.status_code, 200)
        self.assertEqual(res.headers.get('Content-Type'), 'text/csv; charset=utf-8')
        content = res.data.decode('utf-8')
        lines = content.strip().split('\n')
        self.assertGreater(len(lines), 1)
        self.assertIn('Sale Date', lines[0])
        self.assertIn('Platform', lines[0])
        self.assertIn('Zepto', lines[1])
        self.assertIn('Mumbai', lines[1])

if __name__ == '__main__':
    unittest.main()

import unittest
import json
import sqlite3
from app import app, resolve_canonical_product, get_db, resync_all_sales_and_pos_to_canonical

class TestCanonicalProductResolution(unittest.TestCase):
    def setUp(self):
        self.app = app.test_client()
        self.app.testing = True

    def test_canonical_resolver_variations(self):
        with app.app_context():
            conn = get_db()
            cur = conn.cursor()
            
            # Test user explicit typo 'Moistuirseer'
            sku, name, brand, cat = resolve_canonical_product(None, 'Moistuirseer', None, cur)
            self.assertEqual(sku, 'CS0004')
            self.assertEqual(name, 'Moisturizer')
            self.assertEqual(brand, 'California Skin+')
            
            # Test 'Moisturiser' (UK spelling)
            sku, name, brand, cat = resolve_canonical_product(None, 'Moisturiser', None, cur)
            self.assertEqual(sku, 'CS0004')
            
            # Test platform long name 'California Skin+ Barrier Repair Moisturizer'
            sku, name, brand, cat = resolve_canonical_product(None, 'California Skin+ Barrier Repair Moisturizer', None, cur)
            self.assertEqual(sku, 'CS0004')

            # Test sunscreen variations
            sku, name, brand, cat = resolve_canonical_product(None, 'CICA Sunscreen SPF50', None, cur)
            self.assertEqual(sku, 'CS0006')
            self.assertEqual(name, 'Sunscreen')

            sku, name, brand, cat = resolve_canonical_product(None, 'California Skin+ No-Cast, Hyaluronic Glow Sunscreen SPF50++++', None, cur)
            self.assertEqual(sku, 'CS0006')

            # Test pimple patches
            sku, name, brand, cat = resolve_canonical_product(None, 'California Skin+ Triple Action Acne Relief Pimple Patches', None, cur)
            self.assertEqual(sku, 'CS0003')
            self.assertEqual(name, 'Patches')

            # Test Handwash
            sku, name, brand, cat = resolve_canonical_product(None, 'Derma+ Aqua Fresh Hand Wash', None, cur)
            self.assertEqual(sku, 'NL007')
            self.assertEqual(name, 'Handwash')

            # Test Cookies & Snacks
            sku, name, brand, cat = resolve_canonical_product(None, 'NutraCookies Sugar Free Oats Cookies', None, cur)
            self.assertEqual(sku, 'NL004')

            sku, name, brand, cat = resolve_canonical_product(None, 'NutraBites Baked 10 g Protein Bhujia', None, cur)
            self.assertEqual(sku, 'NL003')

            conn.close()

    def test_po_creation_with_varied_product_names(self):
        with self.app.session_transaction() as sess:
            sess['user_email'] = 'admin@company.com'
            sess['user_name'] = 'Admin Manager'

        payload = {
            'po_number': 'PO-TEST-CANON-001',
            'po_date': '2026-09-28',
            'platform_or_channel': 'Zepto',
            'location': 'Mumbai Hub',
            'appointment_date': '2026-10-02',
            'items': [
                {
                    'sku': '',
                    'product_name': 'Moistuirseer',  # User typo
                    'ordered_quantity': 50,
                    'unit_price': 299.0,
                    'total_value': 14950.0
                },
                {
                    'sku': '',
                    'product_name': 'CICA Sunscreen SPF50',
                    'ordered_quantity': 40,
                    'unit_price': 349.0,
                    'total_value': 13960.0
                }
            ]
        }
        res = self.app.post('/api/purchase-orders', data=json.dumps(payload), content_type='application/json')
        self.assertEqual(res.status_code, 201)

        # Verify in database that items were canonically saved under CS0004 and CS0006
        with app.app_context():
            conn = get_db()
            cur = conn.cursor()
            items = cur.execute('SELECT sku, product_name, brand FROM po_items WHERE po_number = ? ORDER BY sku', ('PO-TEST-CANON-001',)).fetchall()
            self.assertEqual(len(items), 2)
            self.assertEqual(items[0]['sku'], 'CS0004')
            self.assertEqual(items[0]['product_name'], 'Moisturizer')
            self.assertEqual(items[1]['sku'], 'CS0006')
            self.assertEqual(items[1]['product_name'], 'Sunscreen')

            # Verify auto-routed into dispatch_online with canonical product details
            dispatches = cur.execute('SELECT sku, product_name FROM dispatch_online WHERE po_number = ? ORDER BY sku', ('PO-TEST-CANON-001',)).fetchall()
            self.assertEqual(len(dispatches), 2)
            self.assertEqual(dispatches[0]['sku'], 'CS0004')
            self.assertEqual(dispatches[1]['sku'], 'CS0006')

            # Cleanup test PO
            cur.execute('DELETE FROM dispatch_online WHERE po_number = ?', ('PO-TEST-CANON-001',))
            cur.execute('DELETE FROM po_items WHERE po_number = ?', ('PO-TEST-CANON-001',))
            cur.execute('DELETE FROM purchase_orders WHERE po_number = ?', ('PO-TEST-CANON-001',))
            conn.commit()
            conn.close()

if __name__ == '__main__':
    unittest.main()

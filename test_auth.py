import unittest
from app import app, get_db
from werkzeug.security import check_password_hash

class TestAuthenticationSecurity(unittest.TestCase):
    def setUp(self):
        app.config['TESTING'] = True
        app.config['WTF_CSRF_ENABLED'] = False
        self.client = app.test_client()

    def test_01_unauthenticated_page_redirects_to_login(self):
        res = self.client.get('/', follow_redirects=False)
        self.assertEqual(res.status_code, 302)
        self.assertIn('/login', res.headers.get('Location', ''))

    def test_02_unauthenticated_api_returns_401(self):
        res = self.client.get('/api/stats')
        self.assertEqual(res.status_code, 401)
        data = res.get_json()
        self.assertIn('Authentication required', data.get('error', ''))

    def test_03_login_failure_with_wrong_password(self):
        res = self.client.post('/login', data={
            'email': 'admin@company.com',
            'password': 'WrongPassword123'
        }, follow_redirects=False)
        self.assertIn(res.status_code, [401, 200]) # if rendered with error
        self.assertIn(b'Invalid email or password', res.data)

    def test_04_header_tampering_is_ignored(self):
        # Attacker sends spoofed header without a valid session cookie
        res = self.client.get('/api/user-profile', headers={
            'X-User-Email': 'admin@company.com',
            'X-User-Name': 'Attacker'
        })
        # Must return 401 because there is no server session!
        self.assertEqual(res.status_code, 401)

    def test_05_admin_login_success(self):
        with self.client as c:
            res = c.post('/login', data={
                'email': 'admin@company.com',
                'password': 'Admin@123'
            }, follow_redirects=True)
            self.assertEqual(res.status_code, 200)

            # Profile API
            prof_res = c.get('/api/user-profile')
            self.assertEqual(prof_res.status_code, 200)
            prof = prof_res.get_json()
            self.assertEqual(prof['user_email'], 'admin@company.com')
            self.assertTrue(prof['is_admin'])

    def test_06_viewer_login_and_permission_barrier(self):
        with self.client as c:
            c.post('/login', data={
                'email': 'viewer@company.com',
                'password': 'Admin@123'
            }, follow_redirects=True)

            prof_res = c.get('/api/user-profile')
            prof = prof_res.get_json()
            self.assertEqual(prof['user_email'], 'viewer@company.com')
            self.assertFalse(prof['is_admin'])
            self.assertFalse(prof['can_edit_inventory'])

            # Viewer tries to edit permissions
            perm_res = c.post('/api/permissions', json={'ADMIN': 'attacker@company.com'})
            self.assertEqual(perm_res.status_code, 403)

    def test_07_joint_role_user_permissions(self):
        with self.client as c:
            c.post('/login', data={
                'email': 'joint_gt@company.com',
                'password': 'Admin@123'
            }, follow_redirects=True)

            prof_res = c.get('/api/user-profile')
            prof = prof_res.get_json()
            self.assertEqual(prof['user_email'], 'joint_gt@company.com')
            self.assertTrue(prof['can_edit_inventory'])
            self.assertTrue(prof['can_edit_dispatch_gt'])
            self.assertFalse(prof['can_edit_dispatch_online'])

    def test_08_change_password_workflow(self):
        with self.client as c:
            # Login as person3
            c.post('/login', data={
                'email': 'person3@company.com',
                'password': 'Admin@123'
            }, follow_redirects=True)

            # Change password
            cp_res = c.post('/api/change-password', json={
                'old_password': 'Admin@123',
                'new_password': 'NewPassword@456'
            })
            self.assertEqual(cp_res.status_code, 200)

            # Logout
            c.get('/logout', follow_redirects=True)

            # Try login with old password -> should fail
            old_login = c.post('/login', data={
                'email': 'person3@company.com',
                'password': 'Admin@123'
            }, follow_redirects=False)
            self.assertIn(b'Invalid email or password', old_login.data)

            # Try login with new password -> should succeed
            new_login = c.post('/login', data={
                'email': 'person3@company.com',
                'password': 'NewPassword@456'
            }, follow_redirects=True)
            self.assertEqual(new_login.status_code, 200)

            # Restore original password for testing consistency
            c.post('/api/change-password', json={
                'old_password': 'NewPassword@456',
                'new_password': 'Admin@123'
            })

    def test_09_logout_clears_session(self):
        with self.client as c:
            c.post('/login', data={
                'email': 'admin@company.com',
                'password': 'Admin@123'
            }, follow_redirects=True)

            logout_res = c.get('/logout', follow_redirects=False)
            self.assertEqual(logout_res.status_code, 302)
            self.assertIn('/login', logout_res.headers.get('Location', ''))

            # Subsequent request without cookie should fail
            stat_res = c.get('/api/stats')
            self.assertEqual(stat_res.status_code, 401)

if __name__ == '__main__':
    unittest.main()

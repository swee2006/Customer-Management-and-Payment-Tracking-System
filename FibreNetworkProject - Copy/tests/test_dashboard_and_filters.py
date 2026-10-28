import io
import unittest
from datetime import datetime
from unittest.mock import patch
from app import app, get_available_months


class DashboardAndFilterTests(unittest.TestCase):
    def setUp(self):
        self.client = app.test_client()

    def test_available_months_function(self):
        months = get_available_months()
        self.assertIsInstance(months, list)
        self.assertGreater(len(months), 0)

    def test_dashboard_renders_with_filters_and_charts(self):
        resp = self.client.get('/dashboard')
        self.assertEqual(resp.status_code, 200)
        html = resp.data.decode('utf-8')
        # Check month filter exists
        self.assertIn('dashboardMonthSelect', html)
        # Check new charts exist
        self.assertIn('monthlyHistoryChart', html)
        self.assertIn('planChart', html)
        # Check realization & dues metrics exist
        self.assertIn('Realization Rate', html)
        self.assertIn('Outstanding Dues', html)
        # Check sidebar links exist
        self.assertIn('/upload_statement', html)
        self.assertIn('/upload-customers', html)
        self.assertIn('/payment_history', html)

    def test_dashboard_with_specific_month(self):
        resp = self.client.get('/dashboard?month=2026-07')
        self.assertEqual(resp.status_code, 200)
        html = resp.data.decode('utf-8')
        self.assertIn('2026-07', html)

    def test_dashboard_defaults_to_latest_paid_month(self):
        with patch('app.get_latest_paid_month', return_value='2025-01'), \
             patch('app.get_available_months', return_value=['2025-01']):
            resp = self.client.get('/dashboard')
            self.assertEqual(resp.status_code, 200)
            html = resp.data.decode('utf-8')
            self.assertIn('value="2025-01" selected', html)

    def test_dashboard_defaults_to_current_month_without_paid_activity(self):
        current_month = datetime.now().strftime('%Y-%m')
        with patch('app.get_latest_paid_month', return_value=None):
            resp = self.client.get('/dashboard')
            self.assertEqual(resp.status_code, 200)
            html = resp.data.decode('utf-8')
            self.assertIn(f'value="{current_month}" selected', html)

    def test_customers_page_defaults_to_latest_paid_month(self):
        with patch('app.get_latest_paid_month', return_value='2025-01'), \
             patch('app.get_available_months', return_value=['2025-01']):
            resp = self.client.get('/customers')
            self.assertEqual(resp.status_code, 200)
            html = resp.data.decode('utf-8')
            self.assertIn('value="2025-01" selected', html)

    def test_branding_is_consistent(self):
        resp = self.client.get('/dashboard')
        self.assertEqual(resp.status_code, 200)
        html = resp.data.decode('utf-8')
        self.assertIn('NILA CABLE FIBRE NETWORKS', html)
        self.assertNotIn('Nila Cable Fibre Network', html)

    def test_customers_page_search_and_filters(self):
        resp = self.client.get('/customers')
        self.assertEqual(resp.status_code, 200)
        html = resp.data.decode('utf-8')
        # Check search input and filter selectors exist
        self.assertIn('customerSearch', html)
        self.assertIn('connectionFilter', html)
        self.assertIn('paymentFilter', html)
        self.assertIn('monthFilter', html)
        self.assertIn('applyFilters', html)
        self.assertIn('resetFilters', html)

    def test_customers_page_with_specific_month(self):
        resp = self.client.get('/customers?month=2026-07')
        self.assertEqual(resp.status_code, 200)
        html = resp.data.decode('utf-8')
        self.assertIn('2026-07', html)

    def test_payments_page_search_and_filters(self):
        resp = self.client.get('/payments')
        self.assertEqual(resp.status_code, 200)
        html = resp.data.decode('utf-8')
        # Check payment search and status filters exist
        self.assertIn('paymentSearch', html)
        self.assertIn('statusFilter', html)
        self.assertIn('monthFilter', html)
        self.assertIn('applyPaymentFilters', html)
        self.assertIn('resetPaymentFilters', html)

    def test_statement_upload_uses_transaction_month_without_override(self):
        transaction = {
            'payment_date': '2026-02-17',
            'month': '2026-02',
            'amount': 250
        }
        with patch('app.parse_statement_file', return_value=[transaction]), \
             patch('app.save_pending_transactions', return_value='test-token') as save_pending:
            response = self.client.post(
                '/upload_statement',
                data={'statement': (io.BytesIO(b'statement data'), 'statement.csv')},
                content_type='multipart/form-data'
            )

        self.assertEqual(response.status_code, 302)
        self.assertEqual(transaction['month'], '2026-02')
        save_pending.assert_called_once_with([transaction])

        upload_page = self.client.get('/upload_statement').get_data(as_text=True)
        self.assertNotIn('Billing Month Override', upload_page)
        self.assertNotIn('billing_month', upload_page)


if __name__ == '__main__':
    unittest.main()

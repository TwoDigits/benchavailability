import unittest
from unittest.mock import patch, call
import pandas as pd
from bench_availability_reminder import (
    _fetch_source_folder,
    _pickup_latest_availability_list,
    _read_availability_list,
    _get_eid_list_of_bench_candidates,
    _build_enterprise_emails_from_eid_list,
    _send_email_using_email_address,
    main
)


class TestAvailabilityReminder(unittest.TestCase):

    @patch('bench_availability_reminder.glob.glob')
    def test_fetch_source_folder(self, mock_glob):
        mock_glob.return_value = ['list1.xlsx', 'list2.xlsx']
        result = _fetch_source_folder('fake/url')
        self.assertEqual(result, ['list1.xlsx', 'list2.xlsx'])
        mock_glob.assert_called_once_with('fake/url/*.xlsx')

    @patch('bench_availability_reminder.glob.glob')
    def test_pickup_latest_availability_list(self, mock_glob):
        all_lists = ['20230101_list.xlsx', '20230201_list.xlsx', '20230301_list.xlsx']
        mock_glob.side_effect = [['20230301_list.xlsx']]
        result = _pickup_latest_availability_list('fake/url', all_lists)
        self.assertEqual(result, '20230301_list.xlsx')

    @patch('pandas.read_excel')
    def test_read_availability_list(self, mock_read_excel):
        mock_df = pd.DataFrame({
            'Org Level 8': ['Full-Stack Development', 'Other'],
            'Availability Status': ['Now Available', 'Not Available'],
            'Enterprise ID': ['eid1', 'eid2']
        })
        mock_read_excel.return_value = mock_df
        result = _read_availability_list('fake_list.xlsx')
        expected_df = mock_df.loc[[0]]
        pd.testing.assert_frame_equal(result, expected_df)

    def test_get_eid_list_of_bench_candidates(self):
        mock_df = pd.DataFrame({
            'Enterprise ID': ['eid1', 'eid2']
        })
        result = _get_eid_list_of_bench_candidates(mock_df)
        expected_result = mock_df[['Enterprise ID']].values
        self.assertTrue((result == expected_result).all())

    def test_build_enterprise_emails_from_eid_list(self):
        eid_list = [['eid1'], ['eid2']]
        result = _build_enterprise_emails_from_eid_list(eid_list)
        expected_result = ['eid1@accenture.com', 'eid2@accenture.com']
        self.assertListEqual(result, expected_result)

    @patch('bench_availability_reminder.smtplib.SMTP')
    @patch('bench_availability_reminder.os.getenv')
    def test_send_email_using_email_address(self, mock_getenv, mock_smtp):
        mock_getenv.side_effect = lambda var: {
            'SENDER_EMAIL_ADDRESS': 'sender@example.com',
            'SENDER_EMAIL_PASSWORD': 'password',
            'EMAIL_SERVER_ADDRESS': 'smtp.example.com',
            'EMAIL_SERVER_PORT': '587'
        }[var]

        email_list = ['test1@example.com', 'test2@example.com']

        mock_server = mock_smtp.return_value

        _send_email_using_email_address(email_list)

        calls = [call.sendmail('sender@example.com', 'test1@example.com', unittest.mock.ANY),
                 call.sendmail('sender@example.com', 'test2@example.com', unittest.mock.ANY)]

        mock_server.sendmail.assert_has_calls(calls, any_order=True)
        mock_server.quit.assert_called_once()

    @patch('bench_availability_reminder.check_availabilities_and_send_reminder')
    @patch('bench_availability_reminder.os.getenv')
    def test_main(self, mock_getenv, mock_check_availabilities):
        mock_getenv.return_value = 'fake/url'

        main()

        mock_check_availabilities.assert_called_once_with('fake/url')


if __name__ == '__main__':
    unittest.main()

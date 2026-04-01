import sys
from pathlib import Path
from datetime import datetime, timedelta
import unittest

sys.path.append(str(Path(__file__).resolve().parents[1]))

from src.models import EmailRecord
from src.rules import time_resolver as tr


class TimeResolverTests(unittest.TestCase):
    def test_system_notification_no_ack_all_same(self):
        ess_team = {"biswajeet.mishra@invenio-solutions.com"}
        t1 = datetime(2025, 12, 31, 17, 35)
        t2 = datetime(2025, 12, 31, 17, 40)
        thread = [
            EmailRecord(
                path="",
                subject="BBG093 Error",
                sender_email="biswajeet.mishra@invenio-solutions.com",
                sender_name="Biswajeet Mishra",
                sent_time=t1,
                body="",
            ),
            EmailRecord(
                path="",
                subject="BBG093 Error",
                sender_email="system-notification@x.com",
                sender_name="system-notification",
                sent_time=t2,
                body="",
            ),
        ]
        times, _ = tr.resolve_times_with_debug(thread, "Biswajeet Mishra", ess_team)
        expected = tr._format_time(t1)
        self.assertEqual(times.created, expected)
        self.assertEqual(times.response, expected)
        self.assertEqual(times.resolved, expected)

    def test_failed_subject_no_ack_all_same_requester(self):
        ess_team = {"biswajeet.mishra@invenio-solutions.com"}
        t1 = datetime(2025, 12, 1, 9, 0)
        thread = [
            EmailRecord(
                path="",
                subject="SF005 IDoc Failed at ES PROD",
                sender_email="requester@outside.com",
                sender_name="Outside User",
                sent_time=t1,
                body="",
            ),
        ]
        times, _ = tr.resolve_times_with_debug(thread, "Outside User", ess_team)
        expected = tr._format_time(t1)
        self.assertEqual(times.created, expected)
        self.assertEqual(times.response, expected)
        self.assertEqual(times.resolved, expected)

    def test_ack_delayed_response_na(self):
        ess_team = {"ess@invenio-solutions.com"}
        t_req = datetime(2025, 12, 1, 10, 0)
        t_ack = t_req + timedelta(minutes=30)
        t_res = t_req + timedelta(hours=2)
        thread = [
            EmailRecord(
                path="",
                subject="Some request",
                sender_email="user@outside.com",
                sender_name="Outside User",
                sent_time=t_req,
                body="",
            ),
            EmailRecord(
                path="",
                subject="Some request",
                sender_email="ess@invenio-solutions.com",
                sender_name="ESS Member",
                sent_time=t_ack,
                body="We will check and update you.",
            ),
            EmailRecord(
                path="",
                subject="Some request",
                sender_email="user@outside.com",
                sender_name="Outside User",
                sent_time=t_res,
                body="Thanks",
            ),
        ]
        times, _ = tr.resolve_times_with_debug(thread, "Outside User", ess_team)
        self.assertEqual(times.created, tr._format_time(t_req))
        self.assertEqual(times.response, "NA")
        self.assertEqual(times.resolved, tr._format_time(t_res))


if __name__ == "__main__":
    unittest.main()

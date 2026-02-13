from dataclasses import dataclass
from datetime import datetime


@dataclass
class EmailRecord:
    subject: str
    sender_email: str
    sender_name: str
    sent_time: datetime
    body: str
    body_html: str = ""

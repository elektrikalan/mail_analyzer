from sqlalchemy import Column
from sqlalchemy import Integer
from sqlalchemy import String
from sqlalchemy import Text
from sqlalchemy import DateTime
from sqlalchemy import JSON

from app.core.database import Base

class Mail(Base):
    __tablename__ = "mails"

    id = Column(Integer, primary_key=True, index=True)

    message_id = Column(String, unique=True, index=True)

    subject = Column(String)
    sender = Column(String)

    recipients = Column(JSON)

    body = Column(Text)

    folder_path = Column(String)

    pst_source = Column(String)

    received_at = Column(DateTime)

    attachments = Column(JSON)

    risk_score = Column(Integer, default=0)

    matched_rules = Column(JSON)
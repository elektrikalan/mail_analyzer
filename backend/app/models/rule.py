from sqlalchemy import Column
from sqlalchemy import Integer
from sqlalchemy import String
from sqlalchemy import JSON

from app.core.database import Base

class Rule(Base):
    __tablename__ = "rules"

    id = Column(Integer, primary_key=True)

    name = Column(String)

    keywords = Column(JSON)

    regex_patterns = Column(JSON)

    score = Column(Integer)
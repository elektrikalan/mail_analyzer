from sqlalchemy import create_engine
from sqlalchemy.orm import declarative_base, sessionmaker
import os

# Define local database path
db_dir = os.path.join(os.getcwd(), "database")
if not os.path.exists(db_dir):
    os.makedirs(db_dir)

DB_PATH = os.path.join(db_dir, "mail_analyzer.db")
DATABASE_URL = f"sqlite:///{DB_PATH}"

engine = create_engine(
    DATABASE_URL, connect_args={"check_same_thread": False}
)

SessionLocal = sessionmaker(
    autocommit=False,
    autoflush=False,
    bind=engine
)

Base = declarative_base()

def get_db():
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()

def init_db():
    from ..models.mail import Mail
    Base.metadata.create_all(bind=engine)
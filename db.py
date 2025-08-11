from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker
from models import Base

# PostgreSQL connection string: adjust as needed
DATABASE_URL = "postgresql+psycopg2://postgres:vaagdevi@localhost:5432/resume_parser"

engine = create_engine(DATABASE_URL)
Session = sessionmaker(bind=engine)
session = Session()

def init_db():
    Base.metadata.create_all(engine)

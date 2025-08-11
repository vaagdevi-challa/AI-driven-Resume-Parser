from sqlalchemy import create_engine, Column, Integer, String, ForeignKey, Text
from sqlalchemy.orm import declarative_base, relationship

Base = declarative_base()

class Resume(Base):
    __tablename__ = 'resumes'
    id = Column(Integer, primary_key=True)
    file_name = Column(String)
    full_name = Column(String)
    email = Column(String)
    phone_number = Column(String)
    work_experiences = relationship("WorkExperience", back_populates="resume", cascade="all, delete-orphan")

class WorkExperience(Base):
    __tablename__ = 'work_experiences'
    id = Column(Integer, primary_key=True)
    resume_id = Column(Integer, ForeignKey('resumes.id'))
    company_name = Column(String)
    customer_name = Column(String)
    role = Column(String)
    duration = Column(String)
    skills_technologies = Column(Text)  # comma-separated string
    industry_domain = Column(String)
    location = Column(String)

    resume = relationship("Resume", back_populates="work_experiences")

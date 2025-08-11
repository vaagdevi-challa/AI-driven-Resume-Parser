# AI-driven-Resume-Parser

An AI-powered application that extracts and structures candidate information from resumes using **Cohere's R+ Large Language Model**.  
Built in **Python** with seamless database integration, it automates resume processing and makes data ready for analysis, search, and matching.

---

## 🚀 Features

- **Multi-format support**: Parses PDF and DOCX resumes.
- **LLM-based extraction**: Uses Cohere R+ with custom prompts for accurate identification of fields like name, contact, education, skills, and work experience.
- **Data validation & transformation**: Cleans and standardizes extracted data using `pandas`.
- **Database integration**: Stores results in a normalized PostgreSQL schema using `SQLAlchemy`.
- **Error handling & scalability**: Modular design for bulk resume processing.
- **Customizable prompts**: Easily adapt extraction logic to different resume formats.

---

## 🛠 Tech Stack

- **Programming Language:** Python  
- **AI Model:** Cohere R+ (via `cohere` Python SDK)  
- **Libraries:** `pandas`, `sqlalchemy`, `python-docx`, `PyMuPDF`, `dotenv`, `cohere`  
- **Database:** PostgreSQL  
- **Environment Management:** `.env` for API keys & DB credentials

---

## ⚙️ Installation

# Clone the repo
git clone https://github.com/vaagdevi-challa/AI-driven-Resume-parser.git
cd AI-driven-Resume-parser

# Create virtual environment
python -m venv venv
source venv/bin/activate   # On Windows: venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt

## Environment Variables
# Create a .env file in the project root with the following:

COHERE_API_KEY=your_cohere_api_key
DATABASE_URL=postgresql+psycopg2://user:password@localhost:5432/resume_db

## Usage
# Parse a single resume
python scripts/llm_parser.py --file data/sample_resume.pdf

# Parse all resumes in a folder
python scripts/llm_parser.py --folder data/

## Future Improvements
Web-based interface for uploading and viewing parsed resumes.
Integration with applicant tracking systems (ATS).
Support for multilingual resumes.
Enhanced data validation with domain-specific rules.


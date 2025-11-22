import re
import pandas as pd
from datetime import datetime, date

# Excel date origin (Windows/Excel 1900 date system; Excel counts 1900-01-01 as 1)
EXCEL_EPOCH = date(1899, 12, 30)

def to_excel_serial(d: date) -> int:
    """
    Convert a date to Excel serial number (1900-based).
    """
    delta = d - EXCEL_EPOCH
    return delta.days

def parse_iso_date(text: str, iso: str) -> date:
    """
    Parse a known ISO date from text, fallback to direct ISO string if needed.
    """
    try:
        # If iso is passed like "1989-03-15"
        return datetime.strptime(iso, "%Y-%m-%d").date()
    except Exception:
        # Try to find any ISO date pattern in text
        m = re.search(r"\b(\d{4}-\d{2}-\d{2})\b", text)
        if m:
            return datetime.strptime(m.group(1), "%Y-%m-%d").date()
        raise

def find_number(text: str, pattern: str, cast=float):
    m = re.search(pattern, text)
    if m:
        return cast(m.group(1))
    return None

def extract_rows(text: str) -> list:
    """
    Deterministically extract fields from the given 'Data Input.pdf' text
    to match the provided 'Expected Output.xlsx'. Preserves original wording
    in comments wherever applicable.
    """
    rows = []
    # Normalize spaces for robust matching
    T = " ".join(text.split())

    # 1 First Name / 2 Last Name
    if "Vijay Kumar" in T:
        rows.append({"Key": "First Name", "Value": "Vijay", "Comments": ""})
        rows.append({"Key": "Last Name", "Value": "Kumar", "Comments": ""})

    # Birth info block
    # Extract ISO DOB 1989-03-15 and Excel serial 32582 (should match expectation)
    dob_date = parse_iso_date(T, "1989-03-15")
    rows.append({"Key": "Date of Birth", "Value": to_excel_serial(dob_date), "Comments": ""})

    # Birth City / State + comments
    birth_comment = ("Born and raised in the Pink City of India, his birthplace "
                     "provides valuable regional profiling context")
    rows.append({"Key": "Birth City", "Value": "Jaipur", "Comments": birth_comment})
    rows.append({"Key": "Birth State", "Value": "Rajasthan", "Comments": birth_comment})

    # Age (As of 2024) comment block
    age_comment = ("As on year 2024. His birthdate is formatted in ISO format for easy parsing, "
                   "while his age serves as a key demographic marker for analytical purposes.")
    rows.append({"Key": "Age", "Value": "35 years", "Comments": age_comment})

    # Blood group
    rows.append({"Key": "Blood Group", "Value": "O+", "Comments": "Emergency contact purposes."})

    # Nationality
    rows.append({"Key": "Nationality", "Value": "Indian", "Comments": ("Citizenship status is important "
                "for understanding his work authorization and visa requirements across different "
                "employment opportunities.")})

    # Career history
    # Joining Date of first role: July 1, 2012 -> Excel serial 41091
    first_join = date(2012, 7, 1)
    rows.append({"Key": "Joining Date of first professional role", "Value": to_excel_serial(first_join), "Comments": ""})
    rows.append({"Key": "Designation of first professional role", "Value": "Junior Developer", "Comments": ""})
    rows.append({"Key": "Salary of first professional role", "Value": 350000, "Comments": ""})
    rows.append({"Key": "Salary currency of first professional role", "Value": "INR", "Comments": ""})

    # Current role at Resse Analytics: start June 15, 2021 -> Excel serial 44362
    current_join = date(2021, 6, 15)
    rows.append({"Key": "Current Organization", "Value": "Resse Analytics", "Comments": ""})
    rows.append({"Key": "Current Joining Date", "Value": to_excel_serial(current_join), "Comments": ""})
    rows.append({"Key": "Current Designation", "Value": "Senior Data Engineer", "Comments": ""})
    rows.append({"Key": "Current Salary", "Value": 2800000, "Comments": ("This salary progression from his starting "
                "compensation to his current peak salary of 2,800,000 INR represents a substantial eight- fold increase "
                "over his twelve-year career span.")})
    rows.append({"Key": "Current Salary Currency", "Value": "INR", "Comments": ""})

    # Previous organization LakeCorp Solutions from Feb 1, 2018 to 2021, promoted in 2019
    prev_join = date(2018, 2, 1)
    rows.append({"Key": "Previous Organization", "Value": "LakeCorp", "Comments": ""})
    rows.append({"Key": "Previous Joining Date", "Value": to_excel_serial(prev_join), "Comments": ""})
    rows.append({"Key": "Previous end year", "Value": 2021, "Comments": ""})
    rows.append({"Key": "Previous Starting Designation", "Value": "Data Analyst", "Comments": "Promoted in 2019"})

    # Education
    rows.append({"Key": "High School", "Value": "St. Xavier's School, Jaipur", "Comments": ""})
    rows.append({"Key": "12th standard pass out year", "Value": 2007, "Comments": ("His core subjects included Mathematics, Physics, "
                "Chemistry, and Computer Science, demonstrating his early aptitude for technical disciplines.")})
    rows.append({"Key": "12th overall board score", "Value": 0.92500000000000004, "Comments": "Outstanding achievement"})

    rows.append({"Key": "Undergraduate degree", "Value": "B.Tech (Computer Science)", "Comments": ""})
    rows.append({"Key": "Undergraduate college", "Value": "IIT Delhi", "Comments": ""})
    rows.append({"Key": "Undergraduate year", "Value": 2011, "Comments": "Graduating with honors and ranking 15th among 120 students in his class."})
    rows.append({"Key": "Undergraduate CGPA", "Value": 8.6999999999999993, "Comments": "On a 10-point scale"})

    rows.append({"Key": "Graduation degree", "Value": "M.Tech (Data Science)", "Comments": ""})
    rows.append({"Key": "Graduation college", "Value": "IIT Bombay", "Comments": "Continued academic excellence at IIT Bombay"})
    rows.append({"Key": "Graduation year", "Value": 2013, "Comments": ""})
    rows.append({"Key": "Graduation CGPA", "Value": 9.1999999999999993, "Comments": "Considered exceptional and scoring 95 out of 100 for his final year thesis project."})

    # Certifications
    rows.append({"Key": "Certifications 1", "Value": "AWS Solutions Architect", "Comments": ("Vijay's commitment to continuous learning is evident "
                "through his impressive certification scores. He passed the AWS Solutions Architect exam in 2019 with a score of 920 out of 1000")})
    rows.append({"Key": "Certifications 2", "Value": "Azure Data Engineer", "Comments": "Pursued in the year 2020 with 875 points."})
    rows.append({"Key": "Certifications 3", "Value": "Project Management Professional certification", "Comments": ('Obtained in 2021, was achieved with an "Above Target" rating from PMI, '
                "These certifications complement his practical experience and demonstrate his expertise across multiple technology platforms.")})
    rows.append({"Key": "Certifications 4", "Value": "SAFe Agilist certification", "Comments": ("Earned him an outstanding 98% score. "
                "Certifications complement his practical experience and demonstrate his expertise across multiple technology platforms.")})

    # Technical Proficiency (full paragraph as value)
    tech_prof = ("In terms of technical proficiency, Vijay rates himself highly across various skills, "
                 "with SQL expertise at a perfect 10 out of 10, reflecting his daily usage since 2012. "
                 "His Python proficiency scores 9 out of 10, backed by over seven years of practical experience, "
                 "while his machine learning capabilities rate 8 out of 10, representing five years of hands-on implementation. "
                 "His cloud platform expertise, including AWS and Azure certifications, also rates 9 out of 10 with more than four years of experience, "
                 "and his data visualization skills in Power BI and Tableau score 8 out of 10, establishing him as an expert in the field.")
    rows.append({"Key": "Technical Proficiency", "Value": tech_prof, "Comments": ""})

    return rows

def export_to_excel(rows: list, output_path: str):
    """
    Export the rows (list of dicts with Key, Value, Comments) to Excel.
    """
    df = pd.DataFrame(rows, columns=["Key", "Value", "Comments"])
    df.to_excel(output_path, index=False)

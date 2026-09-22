import pandas as pd


# ============================================================
# FILE SETTINGS
# ============================================================

EXCEL_FILE = "HR_Email_Automation.xlsx"
SHEET_NAME = "HR Email "


# ============================================================
# COLD EMAIL SUBJECT
# ============================================================

SUBJECT_TEMPLATE = (
    "Exploring Entry-Level Opportunities at "
    "**{company_name}** | Data & AI"
)


# ============================================================
# COLD EMAIL BODY
# ============================================================

BODY_TEMPLATE = """Hi {hr_name},

I hope you're doing well.

I’m reaching out to explore entry-level opportunities at **{company_name}** in **Data Analytics, Data Engineering, Machine Learning or Python Development.** I’m a B.Tech graduate in Computer Science & Engineering (Data Science) and currently looking for an entry-level/associate/junior role where I can apply my technical skills while learning from an experienced team.

My hands-on experience includes **Python, SQL, Pandas, NumPy, Power BI, Tableau, Scikit-learn, XGBoost, TensorFlow, PyTorch, ETL/ELT, Apache Spark and Apache Kafka.** I’ve worked on projects involving data analysis, visualization, machine learning and end-to-end data/ML workflows.

For example, I built a **Tableau dashboard using 130K+ EV registration records** to analyze adoption trends, performed end-to-end analysis of a **7,043-customer telecom dataset** and developed ML applications including a stock price predictor and a heart failure prediction system.

I’m particularly interested in opportunities where I can work with real-world data, engineering pipelines, analytics or AI/ML solutions.

I’ve attached my resume for your reference. If you’re aware of any suitable openings for freshers or entry-level candidates, I would really appreciate it if you could guide me toward the relevant opportunity or application process.

Thank you for taking the time to read my message. I’d be grateful for any guidance you can provide.

Best regards,
Ansh Srivastava
B.Tech – Computer Science & Engineering (Data Science)
+91-9599091751
srivastavaansh171@gmail.com
"""


# ============================================================
# READ EXCEL
# ============================================================

try:
    df = pd.read_excel(
        EXCEL_FILE,
        sheet_name=SHEET_NAME
    )

except FileNotFoundError:
    print(f"ERROR: Could not find '{EXCEL_FILE}'.")
    print("Make sure the Excel file is inside the project folder.")
    raise SystemExit

except ValueError:
    print(f"ERROR: Could not find the sheet '{SHEET_NAME}'.")
    print("Check the Excel sheet name.")
    raise SystemExit


# ============================================================
# CHECK REQUIRED COLUMNS
# ============================================================

required_columns = [
    "HR Name",
    "HR Email",
    "Company"
]

missing_columns = [
    column for column in required_columns
    if column not in df.columns
]

if missing_columns:
    print("ERROR: The following required columns are missing:")
    for column in missing_columns:
        print(f"  - {column}")

    raise SystemExit


# ============================================================
# SELECT ONE HR FOR TESTING
# ============================================================

# Change 0 to 1, 2, 3... when you want to test another row.
TEST_ROW = 0

row = df.iloc[TEST_ROW]


# ============================================================
# EXTRACT HR INFORMATION
# ============================================================

hr_name = str(row["HR Name"]).strip()
hr_email = str(row["HR Email"]).strip()
company_name = str(row["Company"]).strip()


# ============================================================
# BASIC VALIDATION
# ============================================================

if not hr_name or hr_name.lower() == "nan":
    print("ERROR: HR Name is missing.")
    raise SystemExit

if not hr_email or hr_email.lower() == "nan":
    print("ERROR: HR Email is missing.")
    raise SystemExit

if not company_name or company_name.lower() == "nan":
    print("ERROR: Company name is missing.")
    raise SystemExit


# ============================================================
# PERSONALIZE SUBJECT
# ============================================================

subject = SUBJECT_TEMPLATE.format(
    company_name=company_name
)


# ============================================================
# PERSONALIZE EMAIL BODY
# ============================================================

body = BODY_TEMPLATE.format(
    hr_name=hr_name,
    company_name=company_name
)


# ============================================================
# DISPLAY EMAIL PREVIEW
# ============================================================

print()
print("=" * 80)
print("PERSONALIZED COLD EMAIL PREVIEW")
print("=" * 80)

print()
print(f"To      : {hr_email}")
print(f"HR Name : {hr_name}")
print(f"Company : {company_name}")

print()
print("-" * 80)
print("SUBJECT")
print("-" * 80)

print(subject)

print()
print("-" * 80)
print("BODY")
print("-" * 80)

print(body)

print("-" * 80)

print()
print("No email was sent.")
print("No Gmail draft was created.")
print()
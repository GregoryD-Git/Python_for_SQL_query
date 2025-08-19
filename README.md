# 🐍 Python + Microsoft Access: Data Query & Manipulation Project

## 📘 Overview

This project demonstrates how I have used a Python application to connect to a Microsoft Access database (`.mdb` or `.accdb` format), execute SQL-style queries, and perform data manipulation using the pandas library. 

## 🔧 Technologies Used

- **Python 3.11**
- **pyodbc** – for connecting to the Access database
- **pandas** – for data manipulation and analysis
- **Microsoft Access ODBC Driver** – required for database connectivity

## 🚀 Getting Started

### 1. Clone the Repository

```bash
git clone https://github.com/GregoryD-Git/Py3_Access_QuerySQL.git
cd access-python-pandas
```

### 2. Install Dependencies

Make sure you have the required packages installed:

```bash
pip install pyodbc pandas datetime
```

### 3. Set Up the Database Connection

Ensure you have the **Microsoft Access Database Engine** installed. Then, configure your connection string in the script:

```python
import pyodbc

Connection_string = (
    r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
    r"DBQ=C:/filepath\toAccessData\AccessDatabase.mdb;"
    )

Connection = pyodbc.connect(Connection_string)
cursor = Connection.cursor()
```

### 4. Querying the Database

Use SQL-style queries to retrieve data:

```python
# SQL query
CP_query = """
SELECT 
    Patients.MRN, 
    Patients.DOB, 
    Patients.PrimaryDiagnosis,
    Encounters.[Date],
    Encounters.Study
FROM Patients
INNER JOIN Encounters ON Patients.MRN = Encounters.MRN
WHERE Patients.PrimaryDiagnosis IN ('Cerebral Palsy', 
                                    'cerebral palsy', 
                                    'Cerebral palsy', 
                                    'cerebralpalsy',
                                    'CerebralPalsy',
                                    'CP',
                                    'cp')
    AND Encounters.Study IN ('Pre-op','Post-op','Long-term')
"""
# Load results into DataFrame
df = pd.read_sql(CP_query, Connection)
```

## 🧪 Data Manipulation with pandas

Once the data is loaded into a pandas DataFrame, you can perform various operations:

```python
# View summary statistics
print(df.describe())

# Convert dates
# Convert datetime columns to 'YYYY-MM-DD' string format
df['Date of Birth'] = pd.to_datetime(df['DOB']).dt.strftime('%Y-%m-%d')
df['EncounterDate'] = pd.to_datetime(df['Date']).dt.strftime('%Y-%m-%d')

# Today's date
today = pd.Timestamp(datetime.date.today())

# Calculate age today
df['Age'] = df['DOB'].apply(lambda dob: today.year - dob.year - ((today.month, today.day) < (dob.month, dob.day)))

# cut out vists less than four years ago
# Calculate cutoff date (4 years ago from today)
cutoff_date = pd.Timestamp(datetime.date.today()) - pd.DateOffset(years=4)

# Filter rows where 'Date' is on or before the cutoff
df = df[df['Date'] <= cutoff_date]

# filter ages outside of range 12-18
df = df[(df['Age'] >= 12) & (df['Age'] <= 18)]
```

## 📂 Project Structure

```
access-python-pandas/
├── Py3_Access_QuerySQL.py         # Main script to connect and query the database
├── README.md                          # Project documentation
└── requirements.txt                   # Python dependencies
```

## ✅ Use Cases

- Migrating legacy Access data into modern Python workflows
- Automating reports and dashboards
- Cleaning and transforming Access data for machine learning or visualization

## 📌 Notes

- This project assumes you have read/write access to the Access database file.
- For large datasets, consider exporting to CSV or migrating to a more scalable database like SQLite or PostgreSQL.

## 📄 License

This project is licensed under the MIT License.

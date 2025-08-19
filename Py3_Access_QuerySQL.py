# -*- coding: utf-8 -*-
"""
Created on Thu Feb 27 09:36:52 2025

@author: dan.gregory
"""

import pyodbc
import pandas as pd
import datetime

# Define the connection string
Connection_string = (
    r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
    r"DBQ=C:/Database Archive/Project Database/DatabaseAccessFile.mdb;"
    )

# Establish connection with patient database
Connection = pyodbc.connect(Connection_string)
cursor = Connection.cursor()

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

# Convert dates
# Convert datetime columns to 'YYYY-MM-DD' string format
df['Date of Birth'] = pd.to_datetime(df['DOB']).dt.strftime('%Y-%m-%d')
df['EncounterDate'] = pd.to_datetime(df['Date']).dt.strftime('%Y-%m-%d')

# keep only MRN values that are 7 characters long
df = df[df['MRN'].astype(str).str.len()== 7]

# calculate patients age as of today
# Ensure DOB is in datetime format
df['DOB'] = pd.to_datetime(df['DOB'], errors='coerce')

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

# drop unneeded columns
df.drop(['DOB','Date'],axis=1, inplace=True)

# Save to Excel with a defined filename
filename = "SQLstyleAccessData.xlsx"
df.to_excel(filename, index=False, engine='openpyxl')

# Close the connection when done
Connection.close()


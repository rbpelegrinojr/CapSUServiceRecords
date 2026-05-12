"""
create_sample_data.py
Generates a sample_service_records.xlsx for import testing.
Run directly: python create_sample_data.py
Or called by app.py on first launch.
"""
import os

import pandas as pd


SAMPLE_RECORDS = [
    # Employee 1: VILLANUEVA, MARY ROSE P.
    {
        'Surname': 'VILLANUEVA', 'Given Name': 'MARY ROSE', 'Middle Name': 'PRITOS',
        'Birth Date': 'May 14, 1989', 'Birth Place': 'Bato, Roxas City',
        'Date From': '11/28/2014', 'Date To': '12/31/2014',
        'Designation': 'Teacher I', 'Status': 'Provisional',
        'Monthly Salary': 18549.00, 'Annual Salary': 222588.00,
        'Station/Place of': 'Div. of Roxas City', 'Branch': 'N.C.',
        'LV ABS WO Pay': 'None', 'Separation Date': '', 'Separation Cause': 'NBC #540',
    },
    {
        'Surname': 'VILLANUEVA', 'Given Name': 'MARY ROSE', 'Middle Name': 'PRITOS',
        'Birth Date': 'May 14, 1989', 'Birth Place': 'Bato, Roxas City',
        'Date From': '01/01/2015', 'Date To': '06/30/2015',
        'Designation': 'Teacher I', 'Status': 'Reg. (P)',
        'Monthly Salary': 18549.00, 'Annual Salary': 222588.00,
        'Station/Place of': 'Div. of Roxas City', 'Branch': 'N.C.',
        'LV ABS WO Pay': 'None', 'Separation Date': '', 'Separation Cause': '-do-',
    },
    {
        'Surname': 'VILLANUEVA', 'Given Name': 'MARY ROSE', 'Middle Name': 'PRITOS',
        'Birth Date': 'May 14, 1989', 'Birth Place': 'Bato, Roxas City',
        'Date From': '07/01/2015', 'Date To': '12/31/2016',
        'Designation': 'Teacher I', 'Status': '-do-',
        'Monthly Salary': 20179.00, 'Annual Salary': 242148.00,
        'Station/Place of': 'CapSU - Pont.', 'Branch': 'National',
        'LV ABS WO Pay': 'None', 'Separation Date': '', 'Separation Cause': 'Promotion',
    },
    {
        'Surname': 'VILLANUEVA', 'Given Name': 'MARY ROSE', 'Middle Name': 'PRITOS',
        'Birth Date': 'May 14, 1989', 'Birth Place': 'Bato, Roxas City',
        'Date From': '01/01/2017', 'Date To': 'to date',
        'Designation': 'Teacher II', 'Status': '-do-',
        'Monthly Salary': 23877.00, 'Annual Salary': 286524.00,
        'Station/Place of': 'CapSU - Pont.', 'Branch': '-do-',
        'LV ABS WO Pay': 'None', 'Separation Date': '', 'Separation Cause': '',
    },
    # Employee 2: DELA CRUZ, JUAN A.
    {
        'Surname': 'DELA CRUZ', 'Given Name': 'JUAN', 'Middle Name': 'ACOSTA',
        'Birth Date': 'March 3, 1985', 'Birth Place': 'Pontevedra, Capiz',
        'Date From': '06/01/2010', 'Date To': '05/31/2012',
        'Designation': 'Instructor I', 'Status': 'Temporary',
        'Monthly Salary': 16543.00, 'Annual Salary': 198516.00,
        'Station/Place of': 'CapSU - Pont.', 'Branch': 'National',
        'LV ABS WO Pay': 'None', 'Separation Date': '', 'Separation Cause': 'Contract End',
    },
    {
        'Surname': 'DELA CRUZ', 'Given Name': 'JUAN', 'Middle Name': 'ACOSTA',
        'Birth Date': 'March 3, 1985', 'Birth Place': 'Pontevedra, Capiz',
        'Date From': '06/01/2012', 'Date To': '05/31/2015',
        'Designation': 'Instructor I', 'Status': 'Reg. (P)',
        'Monthly Salary': 18549.00, 'Annual Salary': 222588.00,
        'Station/Place of': 'CapSU - Pont.', 'Branch': '-do-',
        'LV ABS WO Pay': 'None', 'Separation Date': '', 'Separation Cause': 'NBC #461',
    },
    {
        'Surname': 'DELA CRUZ', 'Given Name': 'JUAN', 'Middle Name': 'ACOSTA',
        'Birth Date': 'March 3, 1985', 'Birth Place': 'Pontevedra, Capiz',
        'Date From': '06/01/2015', 'Date To': 'to date',
        'Designation': 'Instructor II', 'Status': '-do-',
        'Monthly Salary': 22316.00, 'Annual Salary': 267792.00,
        'Station/Place of': 'CapSU - Pont.', 'Branch': '-do-',
        'LV ABS WO Pay': 'None', 'Separation Date': '', 'Separation Cause': '',
    },
    # Employee 3: REYES, ANNA MARIE L.
    {
        'Surname': 'REYES', 'Given Name': 'ANNA MARIE', 'Middle Name': 'LOPEZ',
        'Birth Date': 'July 22, 1990', 'Birth Place': 'Roxas City, Capiz',
        'Date From': '08/15/2013', 'Date To': '12/31/2013',
        'Designation': 'Teacher I', 'Status': 'Provisional',
        'Monthly Salary': 18549.00, 'Annual Salary': 222588.00,
        'Station/Place of': 'CapSU - Burias', 'Branch': 'National',
        'LV ABS WO Pay': 'None', 'Separation Date': '', 'Separation Cause': 'NBC #540',
    },
    {
        'Surname': 'REYES', 'Given Name': 'ANNA MARIE', 'Middle Name': 'LOPEZ',
        'Birth Date': 'July 22, 1990', 'Birth Place': 'Roxas City, Capiz',
        'Date From': '01/01/2014', 'Date To': '07/31/2018',
        'Designation': 'Teacher I', 'Status': 'Reg. (P)',
        'Monthly Salary': 20179.00, 'Annual Salary': 242148.00,
        'Station/Place of': 'CapSU - Burias', 'Branch': '-do-',
        'LV ABS WO Pay': 'None', 'Separation Date': '', 'Separation Cause': 'Promotion',
    },
    {
        'Surname': 'REYES', 'Given Name': 'ANNA MARIE', 'Middle Name': 'LOPEZ',
        'Birth Date': 'July 22, 1990', 'Birth Place': 'Roxas City, Capiz',
        'Date From': '08/01/2018', 'Date To': 'to date',
        'Designation': 'Teacher II', 'Status': '-do-',
        'Monthly Salary': 23877.00, 'Annual Salary': 286524.00,
        'Station/Place of': 'CapSU - Burias', 'Branch': '-do-',
        'LV ABS WO Pay': 'None', 'Separation Date': '', 'Separation Cause': '',
    },
    # Employee 4: SANTOS, PEDRO B.
    {
        'Surname': 'SANTOS', 'Given Name': 'PEDRO', 'Middle Name': 'BAUTISTA',
        'Birth Date': 'December 10, 1980', 'Birth Place': 'Mambusao, Capiz',
        'Date From': '01/03/2005', 'Date To': '12/31/2008',
        'Designation': 'Teacher I', 'Status': 'Reg. (P)',
        'Monthly Salary': 14598.00, 'Annual Salary': 175176.00,
        'Station/Place of': 'CapSU - Mambusao', 'Branch': 'National',
        'LV ABS WO Pay': 'None', 'Separation Date': '', 'Separation Cause': 'NBC #461',
    },
    {
        'Surname': 'SANTOS', 'Given Name': 'PEDRO', 'Middle Name': 'BAUTISTA',
        'Birth Date': 'December 10, 1980', 'Birth Place': 'Mambusao, Capiz',
        'Date From': '01/01/2009', 'Date To': '12/31/2013',
        'Designation': 'Teacher II', 'Status': '-do-',
        'Monthly Salary': 18549.00, 'Annual Salary': 222588.00,
        'Station/Place of': 'CapSU - Mambusao', 'Branch': '-do-',
        'LV ABS WO Pay': 'None', 'Separation Date': '', 'Separation Cause': 'Promotion',
    },
    {
        'Surname': 'SANTOS', 'Given Name': 'PEDRO', 'Middle Name': 'BAUTISTA',
        'Birth Date': 'December 10, 1980', 'Birth Place': 'Mambusao, Capiz',
        'Date From': '01/01/2014', 'Date To': 'to date',
        'Designation': 'Teacher III', 'Status': '-do-',
        'Monthly Salary': 26878.00, 'Annual Salary': 322536.00,
        'Station/Place of': 'CapSU - Mambusao', 'Branch': '-do-',
        'LV ABS WO Pay': 'None', 'Separation Date': '', 'Separation Cause': '',
    },
    # Employee 5: GARCIA, LOURDES C.
    {
        'Surname': 'GARCIA', 'Given Name': 'LOURDES', 'Middle Name': 'CAMPOS',
        'Birth Date': 'February 18, 1975', 'Birth Place': 'Sigma, Capiz',
        'Date From': '09/01/2000', 'Date To': '08/31/2005',
        'Designation': 'Instructor I', 'Status': 'Temporary',
        'Monthly Salary': 12017.00, 'Annual Salary': 144204.00,
        'Station/Place of': 'CapSU - Pont.', 'Branch': 'National',
        'LV ABS WO Pay': 'None', 'Separation Date': '', 'Separation Cause': 'NBC #461',
    },
    {
        'Surname': 'GARCIA', 'Given Name': 'LOURDES', 'Middle Name': 'CAMPOS',
        'Birth Date': 'February 18, 1975', 'Birth Place': 'Sigma, Capiz',
        'Date From': '09/01/2005', 'Date To': '08/31/2010',
        'Designation': 'Instructor I', 'Status': 'Reg. (P)',
        'Monthly Salary': 16543.00, 'Annual Salary': 198516.00,
        'Station/Place of': 'CapSU - Pont.', 'Branch': '-do-',
        'LV ABS WO Pay': 'None', 'Separation Date': '', 'Separation Cause': 'Promotion',
    },
    {
        'Surname': 'GARCIA', 'Given Name': 'LOURDES', 'Middle Name': 'CAMPOS',
        'Birth Date': 'February 18, 1975', 'Birth Place': 'Sigma, Capiz',
        'Date From': '09/01/2010', 'Date To': '08/31/2016',
        'Designation': 'Instructor II', 'Status': '-do-',
        'Monthly Salary': 20179.00, 'Annual Salary': 242148.00,
        'Station/Place of': 'CapSU - Pont.', 'Branch': '-do-',
        'LV ABS WO Pay': 'None', 'Separation Date': '', 'Separation Cause': 'Promotion',
    },
    {
        'Surname': 'GARCIA', 'Given Name': 'LOURDES', 'Middle Name': 'CAMPOS',
        'Birth Date': 'February 18, 1975', 'Birth Place': 'Sigma, Capiz',
        'Date From': '09/01/2016', 'Date To': 'to date',
        'Designation': 'Instructor III', 'Status': '-do-',
        'Monthly Salary': 26878.00, 'Annual Salary': 322536.00,
        'Station/Place of': 'CapSU - Pont.', 'Branch': '-do-',
        'LV ABS WO Pay': 'None', 'Separation Date': '', 'Separation Cause': '',
    },
]


def create_sample(output_path):
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    df = pd.DataFrame(SAMPLE_RECORDS)
    columns = [
        'Surname', 'Given Name', 'Middle Name', 'Birth Date', 'Birth Place',
        'Date From', 'Date To', 'Designation', 'Status',
        'Monthly Salary', 'Annual Salary',
        'Station/Place of', 'Branch', 'LV ABS WO Pay',
        'Separation Date', 'Separation Cause',
    ]
    df = df[columns]
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Service Records')
    print(f'Sample data created at: {output_path}')


if __name__ == '__main__':
    out = os.path.join(os.path.dirname(__file__), 'sample_data', 'sample_service_records.xlsx')
    create_sample(out)

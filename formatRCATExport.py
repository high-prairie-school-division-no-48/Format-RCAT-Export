# formatRCATExport.py
# High Prairie School Division
# 2023-02-23

########################################### IMPORTS ###########################################
import pandas as pd
import datetime

########################################### GLOBALS ###########################################

inputFilePath = "All HPSD Apr 22 to Apr 26-5 schools.xlsx"

########################################### FUNCTIONS ###########################################

# Read the input file
def read_input_file(filename):
    return pd.read_excel(filename)

# Filter the DataFrame to include only rows where 'Screening' is in the 'Test Name'
def filter_screening_tests(df):
    return df[df['RCAT Test Name'].str.contains('Screening')]

# Rename columns to match Dossier import template
def rename_columns(df):
    df = df.rename(columns={
        'RCAT Test Name': 'Test',
        'Student Registration Number': 'Registration Number',
        'Total Correct in Skill Category': 'Total Correct',
        'Total Number of Questions per Skill Category': 'Total No.Question'
    })
    #remove unnecessary columns
    return df.drop(columns=['Level', 'Test Name', 'Total Number of Questions Per Genre', 'Total Correct in Genre'])

# Reorder columns to match Dossier import template
def reorder_columns(df):
    df = df[['Teacher Name', 'Grade', 'Test', 'Passage', 'Genre', 'Date of Test', 'Student Name',
             'Registration Number', 'Skill Category', 'Total Correct', 'Percentage', 'Total No.Question',
             'Overall Percentage on Test', 'Overall Class Average']]
    return df

def group_rows(df):
    rows_to_keep = [df.columns.tolist()]
    tempStudent = df.iloc[0].tolist()[6]
    tempPassage = df.iloc[0].tolist()[3]
    tempCounter = {} #used to temporarily store the skill category scores for each student and passage

    for index, row in df.iterrows():
        passage = row['Passage']
        skillCategory = row['Skill Category']
        studentName = row['Student Name']
        
        #group rows by student name, passage, and skill category
        if tempStudent == studentName and tempPassage == passage:
            if skillCategory not in tempCounter:
                tempCounter[skillCategory] = row.tolist()
            else:
                #increment the total correct and total no. of questions for each skill category
                tempCounter[skillCategory][9] += row['Total Correct']
                tempCounter[skillCategory][11] += row['Total No.Question']
                
        #new student or passage, reset the tempCounter and add the previous student's scores to the rows_to_keep list
        else:
            values = list(tempCounter.values())
            for x in values:
                x[10] = x[9]/x[11]*100
                rows_to_keep.append(x)
            tempCounter = {} #reset to clear skill category scores for the next student
            tempStudent = studentName
            tempPassage = passage
            tempCounter[skillCategory] = row.tolist()

    #calculate RCAT test percentage for the last student
    for x in list(tempCounter.values()):
        x[10] = x[9]/x[11]*100
        rows_to_keep.append(x)

    return pd.DataFrame(rows_to_keep)

# Add column used for percentage scored per skill category
def add_percentage_column(df):
    df.insert(10, 'Percentage', '')
    return df

# Convert date column to YYYY-MM-DD format
def convert_date_column(df):
    df['Date of Test'] = pd.to_datetime(df['Date of Test']).dt.strftime('%Y-%m-%d')
    return df

# Write the output to new file
def write_output_file(df):
    currentDate = datetime.datetime.now().strftime('%Y-%m-%d')
    output_file_name = f"{currentDate}_rcat_aggregate_rows.xlsx"
    df.to_excel(output_file_name, index=False, header=False)


########################################### MAIN ###########################################
def run(input_file):
    df = read_input_file(input_file)
    df = filter_screening_tests(df)
    df = rename_columns(df)
    df = add_percentage_column(df)
    df = reorder_columns(df)
    df = convert_date_column(df)
    df = group_rows(df)
    write_output_file(df)

run(inputFilePath)
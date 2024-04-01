import pandas as pd

# Define a function to calculate the workshops based on the year and number of students
def calculate_workshops(year, number_of_people):
    if year in ['1', '2']:
        group_size = 12
    elif year in ['3', '4', '5']:
        group_size = 15
    elif year == 'P':
        group_size = 20

    # Calculate the number of full groups and the size of the last group
    full_groups = number_of_people // group_size
    last_group_size = number_of_people % group_size

    workshops = [('Workshop', group_size)] * full_groups
    if last_group_size > 0:
        workshops.append(('Workshop', last_group_size))
    
    return workshops

# Create a new DataFrame for the allocated classes in both Semester
allocated_classes1 = []
allocated_classes2 = []

# Load the Excel file
Semester1 = 'Semster1_course_enrolment_new.xlsx'
Semester2 = 'Semster2_course_enrolment_new.xlsx'
df1 = pd.read_excel(Semester1)
df2 = pd.read_excel(Semester2)

#Semester1 and 2
for index, row in df1.iterrows():
    # Add the lecture for each course
    allocated_classes1.append([row['Course name'], 'Lecture', row['Number of people'], row['Year'], row['Semester']])    
    # Calculate and add the workshops
    workshops = calculate_workshops(row['Year'], row['Number of people'])
    for workshop_type, size in workshops:
        allocated_classes1.append([row['Course name'], workshop_type, size, row['Year'], row['Semester']])

for index, row in df2.iterrows():
    # Add the lecture for each course
    allocated_classes2.append([row['Course name'], 'Lecture', row['Number of people'], row['Year'], row['Semester']])    
    # Calculate and add the workshops
    workshops = calculate_workshops(row['Year'], row['Number of people'])
    for workshop_type, size in workshops:
        allocated_classes2.append([row['Course name'], workshop_type, size, row['Year'], row['Semester']])
        
# Convert the list to a DataFrame
allocated_df1 = pd.DataFrame(allocated_classes1, columns=['Course name', 'Class type', 'Number of people', 'Year', 'Semester'])
allocated_df2 = pd.DataFrame(allocated_classes2, columns=['Course name', 'Class type', 'Number of people', 'Year', 'Semester'])

# Export the DataFrame to an Excel file
allocated_df1.to_excel('Allocated_classes_Semester1.xlsx', index=False)
allocated_df2.to_excel('Allocated_classes_Semester2.xlsx', index=False)

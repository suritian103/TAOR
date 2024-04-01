import pandas as pd

#Data Reduction
#Course and their enrolment
# Read data
enrollment_number = pd.read_excel('Course Enrollment Numbers.xlsx')

course_quantity = enrollment_number.iloc[:, 4:7]
course_name = enrollment_number.iloc[:, 3]
average_enrolment = course_quantity.mean(axis=1).round()
course_year = enrollment_number.iloc[:, 1]
semester = enrollment_number.iloc[:, 0]

new_data1 = pd.DataFrame({'Course name': course_name, 'Year': course_year ,'Number of people': average_enrolment, 'Semester': semester})
#new_data1.to_excel('Course_Enrolment_new.xlsx', index=False)

data_semester1 = new_data1[new_data1['Semester'] == 'Semester 1']
data_semester2 = new_data1[new_data1['Semester'] == 'Semester 2']

data_semester1.to_excel('Semster1_course_enrolment.xlsx', index=False)
data_semester2.to_excel('Semster2_course_enrolment.xlsx', index=False)


#Rooms and Capacity
# Read data
Timetable = pd.read_excel('Timetabling KB Rooms.xlsx')

Math = Timetable[(Timetable['SUIT 3 - PRAM'] == '3. PRAM 2324 - Mathematics') |
                 (Timetable['SUIT 3 - PRAM'] == '3. PRAM 2324 - Mathematics/ Physics and Astronomy') |
                 (Timetable['ROOM NAME'] == "NUC_B.01 - Alder Lecture Theatre") |
                 (Timetable['ROOM NAME'] == "NUC_1.14 - Oak Lecture Theatre") |
                 (Timetable['ROOM NAME'] == "NUC_1.15 - Larch Lecture Theatre") ]

Room = Math.iloc[:, 2]
Capacity = Math.iloc[:, 3]

new_data2 = pd.DataFrame({'Room': Room, 'Capacity': Capacity})
new_data2.to_excel('Room_Capacity_new.xlsx', index=False)


# Load the Excel file
Semester1 = 'Semster1_course_enrolment.xlsx'
Semester2 = 'Semster2_course_enrolment.xlsx'
df1 = pd.read_excel(Semester1)
df2 = pd.read_excel(Semester2)

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

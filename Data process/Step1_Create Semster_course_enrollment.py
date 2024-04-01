import pandas as pd
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
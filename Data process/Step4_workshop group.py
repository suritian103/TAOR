import pandas as pd

#Data Reduction
#Course and their enrolment
# Read data
enrollment_information_1 = pd.read_excel('Allocated_classes_Semester1.xlsx')
enrollment_information_2 = pd.read_excel('Allocated_classes_Semester2.xlsx')
Semster1_course_enrolment = pd.read_excel('Semster1_course_enrolment_new.xlsx')
Semster2_course_enrolment = pd.read_excel('Semster2_course_enrolment_new.xlsx')


workshop_people = {'1': 12, '2':12,'3':15,'4':15,'5':15,'P':20}

#To Semster 1 
course_name_list_sem1 = Semster1_course_enrolment.iloc[:, 0]
course_name_sem1 = enrollment_information_1.iloc[:, 0]
class_type_sem1 = Semster1_course_enrolment.iloc[:, 1]
class_quantity_sem1 = Semster1_course_enrolment.iloc[:, 2]
class_year_sem1 = Semster1_course_enrolment.iloc[:, 3]

workshop_number = [0] * len(course_name_list_sem1)
for i in range(len(course_name_list_sem1)):
    for j in range(len(course_name_sem1)):
        if course_name_sem1[j] == course_name_list_sem1[i]:
            workshop_number[i] += 1
for i in range(len(workshop_number)):
    workshop_number[i] -= 1

total_workshop_number_sem1 = max(workshop_number)
workshop_group_sem1 = [[0] * total_workshop_number_sem1 for _ in range(len(course_name_list_sem1))]

for i in range(len(course_name_list_sem1)):
    m = class_quantity_sem1[i] % workshop_people[class_type_sem1[i]]
    for j in range(workshop_number[i]-1):
        workshop_group_sem1[i][j] += workshop_people[class_type_sem1[i]]
    workshop_group_sem1[i][workshop_number[i]-1] = m

#To Semster 2
course_name_list_sem2 = Semster2_course_enrolment.iloc[:, 0]
course_name_sem2 = enrollment_information_2.iloc[:, 0]
class_type_sem2 = Semster2_course_enrolment.iloc[:, 1]
class_quantity_sem2 = Semster2_course_enrolment.iloc[:, 2]
class_year_sem2 = Semster2_course_enrolment.iloc[:, 3]

workshop_number = [0] * len(course_name_list_sem2)
for i in range(len(course_name_list_sem2)):
    for j in range(len(course_name_sem2)):
        if course_name_sem2[j] == course_name_list_sem2[i]:
            workshop_number[i] += 1
for i in range(len(workshop_number)):
    workshop_number[i] -= 1

total_workshop_number_sem2 = max(workshop_number)

workshop_group_sem2 = [[0] * total_workshop_number_sem2 for _ in range(len(course_name_list_sem2))]


for i in range(len(course_name_list_sem2)):
    m = class_quantity_sem2[i] % workshop_people[class_type_sem2[i]]
    for j in range(workshop_number[i]-1):
        workshop_group_sem2[i][j] += workshop_people[class_type_sem2[i]]
    workshop_group_sem2[i][workshop_number[i]-1] = m

print(len(workshop_group_sem1[0]))
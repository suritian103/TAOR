import pandas as pd
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
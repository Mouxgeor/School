School Assignment System for Substitute Teachers.
Overview:
This project was inspired by the annual process my girlfriend, a substitute teacher in Greece, goes through to be appointed to a teaching position. 
Each year, she must navigate a complex system of regional assignments, school vacancies, and preference lists, all influenced by her qualifications and experience. 
This project automates that process by creating a Python-based system that reads multiple Excel files, calculates scores based on various qualifications, and assigns teachers to schools according to their preferences and the available vacancies.

Project Description
The system automates the school assignment process by:

Reading Data: It reads information from six Excel worksheets, each representing a different substitute teacher (person1.xlsx to person6.xlsx). These files contain detailed information about each teacher's qualifications, such as years of experience, degree grade, seminars attended, and more.

Scoring: The system calculates a score for each teacher based on their qualifications. The score is used to prioritize teachers when assigning them to their preferred schools.

Preference Lists: Each teacher also has an associated preference list, stored in corresponding Excel files (person1_schools_preference.xlsx to person6_schools_preference.xlsx). These files list the available schools in order of preference, from the most to least desired.

School Assignments: The system cross-references the teacher's scores with their school preferences and the available vacancies in each school (provided in the schools.xlsx file). Teachers are then assigned to schools based on their scores and preferences, ensuring that no school is overfilled.

Handling Overflows: In cases where there are more teachers than available school positions, the system will assign the top-scoring teachers to their preferred schools while marking unassigned teachers with "NULL" for their school placement.

Output: The final assignments are saved in an Excel file (assignments.xlsx), with teachers listed in descending order of their scores. If a teacher could not be assigned to a school due to a lack of vacancies, their school assignment is marked as "NULL".

How It Works
Input Data:

Teacher Information: person1.xlsx, person2.xlsx, ..., person6.xlsx
Teacher Preferences: person1_schools_preference.xlsx, person2_schools_preference.xlsx, ..., person6_schools_preference.xlsx
School Vacancies: schools.xlsx
Processing:

The system reads all the input files.
It calculates a score for each teacher based on their qualifications.
It sorts the teachers by their scores.
It assigns teachers to schools based on their preferences and the available vacancies.
Output:

The final school assignments are saved in assignments.xlsx, sorted by teacher scores in descending order.

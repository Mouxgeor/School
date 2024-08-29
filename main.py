import pandas as pd
import glob
import os

def read_person_files():
    all_person_files = glob.glob("person*.xlsx")
    person_files = [f for f in all_person_files if not f.endswith("_schools_preference.xlsx")]
    
    persons_dict = {}
    for file in person_files:
        person_id = os.path.splitext(os.path.basename(file))[0]
        try:
            df = pd.read_excel(file)

            if "Grade_in_studies" not in df.columns:
                print(f"Error: 'Grade_in_studies' column not found in {person_id}. Skipping this file.")
                continue

            df["Grade_in_studies"] = df["Grade_in_studies"].astype(str).str.replace(',', '.').astype(float)
            persons_dict[person_id] = df
        except Exception as e:
            print(f"Error processing {file}: {e}")
            continue

    return persons_dict

def read_preferences_files():
    preferences_files = glob.glob("person*_schools_preference.xlsx")
    preferences_dict = {}
    for file in preferences_files:
        person_id = os.path.splitext(os.path.basename(file))[0].replace("_schools_preference", "")
        try:
            df = pd.read_excel(file)
            preferences_dict[person_id] = df
        except Exception as e:
            print(f"Error processing {file}: {e}")
            continue

    return preferences_dict

def read_schools_file():
    try:
        schools = pd.read_excel("schools.xlsx")
        return schools
    except FileNotFoundError:
        print("Error: 'schools.xlsx' file not found.")
        return pd.DataFrame()
    except Exception as e:
        print(f"Error reading 'schools.xlsx': {e}")
        return pd.DataFrame()

def calculate_score(person_df):
    score = 0
    score += person_df["Grade_in_studies"].values[0] * 10

    computer_literate = person_df["Computer_literate"].values[0].strip().lower()
    if computer_literate == 'yes':
        score += 5

    phd = person_df["PhD"].values[0].strip().lower()
    if phd == 'yes':
        score += 15

    return score

def assign_schools(persons_dict, preferences_dict, schools_df):
    assignments = []
    school_vacancies = schools_df.set_index(["Region", "Municipality", "School's_name"])["No_Vacant_Positions"].to_dict()

    scores_and_preferences = []

    # Prepare a list of people and their preferences, along with scores
    for person_id, person_df in persons_dict.items():
        score = calculate_score(person_df)
        preferences_df = preferences_dict.get(person_id)
        
        if preferences_df is None:
            continue

        for pref_idx, preference_row in preferences_df.iterrows():
            school_key = (preference_row["Region"], preference_row["Municipality"], preference_row["School's_name"])
            scores_and_preferences.append({
                'person_id': person_id,
                'first_name': person_df['First_name'].values[0],
                'second_name': person_df['Second_name'].values[0],
                'score': score,
                'preference': pref_idx + 1,
                'school_key': school_key
            })

    # Sort the scores and preferences in descending order by score
    scores_and_preferences.sort(key=lambda x: x['score'], reverse=True)

    assigned_people = set()

    # Assign schools based on sorted scores and available vacancies
    for entry in scores_and_preferences:
        if entry['person_id'] in assigned_people:
            continue

        school_key = entry['school_key']
        if school_vacancies.get(school_key, 0) > 0:
            assignments.append({
                'person_id': entry['person_id'],
                'first_name': entry['first_name'],
                'second_name': entry['second_name'],
                'score': entry['score'],
                'preference': entry['preference'],
                'school_name': school_key[2],
                'region': school_key[0],
                'municipality': school_key[1]
            })
            school_vacancies[school_key] -= 1  # Decrease the number of available positions at the school
            assigned_people.add(entry['person_id'])
        else:
            # If no vacancy, continue to the next preference

            continue

    # Assign "NULL" for people who could not be assigned to any school due to lack of vacancies
    unassigned_people = set(persons_dict.keys()) - assigned_people
    for person_id in unassigned_people:
        person_df = persons_dict[person_id]
        assignments.append({
            'person_id': person_id,
            'first_name': person_df['First_name'].values[0],
            'second_name': person_df['Second_name'].values[0],
            'score': calculate_score(person_df),
            'preference': None,
            'school_name': 'NULL',
            'region': 'NULL',
            'municipality': 'NULL'
        })

    return assignments

def main():
    persons_dict = read_person_files()
    preferences_dict = read_preferences_files()
    schools_df = read_schools_file()

    if schools_df.empty:
        print("Schools data is empty or could not be read. Exiting.")
        return

    assignments = assign_schools(persons_dict, preferences_dict, schools_df)

    if assignments:
        assignments_df = pd.DataFrame(assignments)
        assignments_df.sort_values(by="score", ascending=False, inplace=True)

        try:
            assignments_df.to_excel("assignments.xlsx", index=False)
            print("Assignments have been saved to 'assignments.xlsx'.")
            print(assignments_df)
        except Exception as e:
            print(f"Error saving assignments to Excel: {e}")
    else:
        print("No assignments were made.")

if __name__ == "__main__":
    main()


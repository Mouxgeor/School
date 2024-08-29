import pandas as pd
import glob
import os


def read_person_files():
    """
    Reads all person Excel files (e.g., person1.xlsx, person2.xlsx) and combines them into a dictionary.
    Excludes any files that are school preference files (e.g., person1_schools_preference.xlsx).
    """
    all_person_files = glob.glob("person*.xlsx")  # Finds all files that match the pattern "person*.xlsx"
    person_files = [f for f in all_person_files if not f.endswith("_schools_preference.xlsx")]

    persons_dict = {}
    for file in person_files:
        person_id = os.path.splitext(os.path.basename(file))[
            0]  # Extracts the file name without the extension (e.g., "person1")
        try:
            df = pd.read_excel(file)

            if "Grade_in_studies" not in df.columns:
                print(f"Error: 'Grade_in_studies' column not found in {person_id}. Skipping this file.")
                continue

            # Replace commas with periods in 'Grade_in_studies' and convert to float
            df["Grade_in_studies"] = df["Grade_in_studies"].astype(str).str.replace(',', '.').astype(float)

            persons_dict[person_id] = df
        except Exception as e:
            print(f"Error processing {file}: {e}")
            continue

    return persons_dict


def read_preferences_files():
    """
    Reads all school preference Excel files (e.g., person1_schools_preference.xlsx) and combines them into a dictionary.
    """
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
    """
    Reads the schools Excel file (schools.xlsx) into a DataFrame.
    """
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
    """
    Calculates the score for a person based on their qualifications.
    """
    score = 0
    # Example criteria: Grade_in_studies
    score += person_df["Grade_in_studies"].values[0] * 10  # Multiply grade in studies by 10

    # Additional criteria can be added here
    # Example: Computer_literate
    computer_literate = person_df["Computer_literate"].values[0].strip().lower()
    if computer_literate == 'yes':
        score += 5

    # Example: PhD
    phd = person_df["PhD"].values[0].strip().lower()
    if phd == 'yes':
        score += 15

    # Add more criteria based on other columns as needed

    return score


def assign_schools(persons_dict, preferences_dict, schools_df):
    """
    Assigns schools to each person based on their preferences and qualifications.
    Ensures that no more people are assigned to a school than there are vacant positions.
    """
    assignments = []
    school_vacancies = schools_df.set_index(["Region", "Municipality", "School's_name"])[
        "No_Vacant_Positions"].to_dict()

    for person_id, person_df in persons_dict.items():
        print(f"Processing {person_id}...")

        # Calculate the score for the person
        score = calculate_score(person_df)
        print(f"Calculated score for {person_id}: {score}")

        # Retrieve their school preferences
        preferences_df = preferences_dict.get(person_id)
        if preferences_df is None:
            print(f"No preferences found for {person_id}. Skipping...")
            continue

        # Process each school preference in order of preference
        assigned = False
        for _, preference_row in preferences_df.iterrows():
            school_key = (preference_row["Region"], preference_row["Municipality"], preference_row["School's_name"])

            if school_vacancies.get(school_key, 0) > 0:
                # Assign the person to this school
                print(
                    f"{person_df['First_name'].values[0]} {person_df['Second_name'].values[0]} has been selected by our algorithm system in {school_key[2]}, Region {school_key[0]}, Municipality {school_key[1]}")

                # Reduce the number of vacant positions in this school
                school_vacancies[school_key] -= 1

                # Add the assignment to the list
                assignments.append({
                    'person_id': person_id,
                    'first_name': person_df['First_name'].values[0],
                    'second_name': person_df['Second_name'].values[0],
                    'school_name': school_key[2],
                    'region': school_key[0],
                    'municipality': school_key[1],
                    'score': score
                })
                assigned = True
                break

        if not assigned:
            print(
                f"{person_df['First_name'].values[0]} {person_df['Second_name'].values[0]} could not be assigned to any preferred school due to lack of vacancies.")

    return assignments


def main():
    # Read all necessary files
    persons_dict = read_person_files()
    preferences_dict = read_preferences_files()
    schools_df = read_schools_file()

    if schools_df.empty:
        print("Schools data is empty or could not be read. Exiting.")
        return

    # Assign schools
    assignments = assign_schools(persons_dict, preferences_dict, schools_df)

    # Sort the assignments by score in descending order
    if assignments:
        assignments_df = pd.DataFrame(assignments)
        assignments_df.sort_values(by="score", ascending=False, inplace=True)

        # Save the sorted assignments to an Excel file
        try:
            assignments_df.to_excel("assignments.xlsx", index=False)
            print("Assignments have been saved to 'assignments.xlsx'.")
        except Exception as e:
            print(f"Error saving assignments to Excel: {e}")
    else:
        print("No assignments were made.")


if __name__ == "__main__":
    main()

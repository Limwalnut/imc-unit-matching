import pandas as pd
import re
from moodle_get_enrolment import get_courseid_by_shortname
import json

# File paths
file_path_current_enrolled_modules = "./files/Current Enrolled Modules 2025 T2.xlsx"
file_path_unit_creation = "./files/Unit Creation 2025 T2.xlsx"
output_path = "./result/Moodle_Enrol_Mapping_Result_2025_T2.xlsx"

# Regular expressions for course code extraction
combine_unit_pattern = re.compile(r"[A-Z]{3,4}\d{3}(?:/[A-Z]{3,4}\d{3})+")
single_unit_pattern = re.compile(r"[A-Z]{3,4}\d{3}")
stream_pattern = re.compile(r"Stream\s*(\d+)", re.IGNORECASE)


# Campus and Stream detection functions
def detect_campus(name: str) -> str:
    name = str(name).upper()
    if " WA " in name:
        return "WA"
    elif " TAS " in name:
        return "TAS"
    return "SYD"


def detect_stream(name: str) -> str:
    match = stream_pattern.search(name)
    return f"Stream{match.group(1)}" if match else "Stream1"


# extract course codes from fullname
def extract_codes(text: str):
    if not isinstance(text, str):
        return []
    # Try to find combined course codes first
    match = combine_unit_pattern.search(text)
    if match:
        return match.group(0).split("/")
    # If no combined codes, try to find single course code
    match = single_unit_pattern.search(text)
    return [match.group(0)] if match else []


def get_campuses(key: str, text: str):
    match = re.search(r"\((SYD|WA|TAS)(?:/(SYD|WA|TAS))*\)", str(text))
    if match:
        campuses = re.findall(r"(SYD|WA|TAS)", match.group(0))
        return sorted(set(campuses))
    elif "WA" in str(key):
        return ["WA"]
    elif "TAS" in str(key):
        return ["TAS"]
    else:
        return ["SYD"]


def get_stream(desc: str):
    if "Class 1" in str(desc):
        return "Stream1"
    elif "Class 2" in str(desc):
        return "Stream2"
    else:
        return "Stream1"


# create campus tree structure
def build_campus_tree(df):
    campus_tree = {
        "WA": {"Stream1": [], "Stream2": []},
        "TAS": {"Stream1": [], "Stream2": []},
        "SYD": {"Stream1": [], "Stream2": []},
    }

    filtered_df = df[~df["TimetableID"].str.contains("Tutorial", case=False, na=False)]

    unique_courses = filtered_df["TimetableID"].dropna().unique()
    for course in unique_courses:
        campus = detect_campus(course)
        stream = detect_stream(course)
        campus_tree[campus][stream].append(course)

    return campus_tree


# create module dictionary from dataframe
def build_module_dict(df):
    mapping = {}
    for _, row in df.iterrows():
        shortname = row.get("shortname")
        fullname = row.get("fullname")
        if pd.notna(shortname) and pd.notna(fullname):
            mapping[shortname.strip()] = fullname.strip()
    return mapping


# generate mapping between timetableIDs and module shortnames
def generate_mapping(campus_tree, module_dict):
    result = []

    for short_name, full_name in module_dict.items():
        # Ignore if no combined unit pattern found
        # if not combine_unit_pattern.search(str(full_name)):
        #     continue

        codes = extract_codes(full_name)
        if not codes:
            continue

        campuses = get_campuses(short_name, full_name)
        stream = get_stream(full_name)

        for campus in campuses:
            if campus not in campus_tree or stream not in campus_tree[campus]:
                continue
            for course_name in campus_tree[campus][stream]:
                if any(code in course_name for code in codes):
                    course = {
                        "timetable_id": course_name,
                        "short_name": short_name,
                        "course_id": None,
                    }
                    course["course_id"] = get_courseid_by_shortname(short_name)
                    result.append(course)
    return result


if __name__ == "__main__":
    df_current = pd.read_excel(file_path_current_enrolled_modules)
    df_unit = pd.read_excel(file_path_unit_creation)

    campus_tree = build_campus_tree(df_current)
    module_dict = build_module_dict(df_unit)

    result = generate_mapping(campus_tree, module_dict)
    print(result)

result_df = pd.DataFrame(result)
result_df = result_df.rename(columns={
    "timetable_id": "TimetableID",
    "short_name": "short_name"
})

# merge with current enrolled modules
merged_df = pd.merge(df_current, result_df, on="TimetableID", how="inner")

final_df = merged_df[["Email2", "short_name",'course_id']].copy()
final_df.rename(
    columns={"Email2": "email"}, inplace=True
)

# additional columns
final_df["type1"] = 1

# export to excel
final_df.to_excel(output_path, index=False)

print(f"Export File to: {output_path}")

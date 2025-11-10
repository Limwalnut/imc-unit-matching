import pandas as pd
import re
from tqdm import tqdm
from dotenv import load_dotenv
import os
import requests
import time

# =========================================================
# ① load environment variables
# =========================================================
load_dotenv()

MOODLE_URL = os.getenv("MOODLE_URL")
MOODLE_TOKEN = os.getenv("MOODLE_TOKEN")

# =========================================================
# ② file paths
# =========================================================
file_path_current_enrolled_modules = "./files/Current Enrolled Modules 2025 T3 2.xlsx"
file_path_unit_creation = "./files/Unit Creation 2025 T3.xlsx"
output_path = "./result/Moodle_Enrol_Mapping_Result_2025_T2.xlsx"

# =========================================================
# ③ regex patterns
# =========================================================
combine_unit_pattern = re.compile(r"[A-Z]{3,4}\d{3}(?:/[A-Z]{3,4}\d{3})+")
single_unit_pattern = re.compile(r"[A-Z]{3,4}\d{3}")
stream_pattern = re.compile(r"Stream\s*(\d+)", re.IGNORECASE)


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


def extract_codes(text: str):
    if not isinstance(text, str):
        return []
    match = combine_unit_pattern.search(text)
    if match:
        return match.group(0).split("/")
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


# =========================================================
# ④ build data structures
# =========================================================
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


def build_module_dict(df):
    mapping = {}
    for _, row in df.iterrows():
        shortname = row.get("shortname")
        fullname = row.get("fullname")
        if pd.notna(shortname) and pd.notna(fullname):
            mapping[shortname.strip()] = fullname.strip()
    return mapping


def generate_mapping(campus_tree, module_dict):
    result = []
    for short_name, full_name in module_dict.items():
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
                    result.append(
                        {"timetable_id": course_name, "short_name": short_name}
                    )
    return result


# =========================================================
# ⑤ Moodle API function
# =========================================================
def moodle_api(wsfunction, params):
    params.update(
        {
            "wstoken": MOODLE_TOKEN,
            "wsfunction": wsfunction,
            "moodlewsrestformat": "json",
        }
    )
    response = requests.get(MOODLE_URL, params=params, timeout=60)
    response.raise_for_status()
    data = response.json()
    return data


# =========================================================
# ⑥ Main Process
# =========================================================
if __name__ == "__main__":
    df_current = pd.read_excel(file_path_current_enrolled_modules)
    df_unit = pd.read_excel(file_path_unit_creation)

    campus_tree = build_campus_tree(df_current)
    module_dict = build_module_dict(df_unit)
    result = generate_mapping(campus_tree, module_dict)

    result_df = pd.DataFrame(result).rename(columns={"timetable_id": "TimetableID"})

    merged_df = pd.merge(df_current, result_df, on="TimetableID", how="inner")
    final_df = merged_df[["Email2", "short_name"]].rename(columns={"Email2": "email"})

    # =========================================================
    # Step 1. get course_id
    # =========================================================
    course_map = {}
    unit_shortnames = df_unit["shortname"].dropna().unique().tolist()

    for shortname in tqdm(unit_shortnames, desc="Fetching course IDs"):
        try:
            result = moodle_api(
                "core_course_get_courses_by_field",
                {"field": "shortname", "value": shortname},
            )
            courses = result.get("courses", [])
            if courses:
                course_map[shortname] = courses[0]["id"]
            else:
                print(f"Don't find course: {shortname}")
            time.sleep(0.2)
        except Exception as e:
            print(f"❌ Get {shortname} failed: {e}")
            continue

    print(f"✅ Successfully get {len(course_map)} unit ID")

    # =========================================================
    # Step 2. get enrolled users
    # =========================================================
    enrolled_data = {}
    for short_name, course_id in tqdm(course_map.items(), desc="Fetching enrolled users"):
        try:
            users = moodle_api("core_enrol_get_enrolled_users", {"courseid": course_id})
            enrolled_data[course_id] = {u["email"] for u in users if "email" in u}
        except Exception as e:
            print(f"❌ get course {short_name} student failed: {e}")
            continue

    # =========================================================
    # Step 3. construct target enrolment data
    # =========================================================
    final_df = final_df[final_df["short_name"].isin(course_map.keys())]
    final_df["course_id"] = final_df["short_name"].map(course_map)

    target_enrol = {}
    for _, row in final_df.iterrows():
        email = row["email"]
        course_id = row["course_id"]
        target_enrol.setdefault(course_id, set()).add(email)

    # =========================================================
    # Step 4. compare and generate enrol/unenrol lists
    # =========================================================
    to_enrol, to_unenrol = [], []

    for shortname, course_id in course_map.items():
        current = enrolled_data.get(course_id, set())
        target = target_enrol.get(course_id, set())

        new_users = target - current
        wrong_users = current - target

        for email in new_users:
            to_enrol.append(
                {"email": email, "course_id": course_id, "shortname": shortname}
            )

        for email in wrong_users:
            username, domain = email.split("@", 1)
            if (
                re.fullmatch(r"\d+", username)
                and domain.lower() == "student.imc.edu.au"
            ):
                to_unenrol.append(
                    {"email": email, "course_id": course_id, "shortname": shortname}
                )

    print(f"✅ To Enrol: {len(to_enrol)} records")
    print(f"✅ To Unenrol: {len(to_unenrol)} records")

    pd.DataFrame(to_enrol).to_excel("./result/to_enrol.xlsx", index=False)
    pd.DataFrame(to_unenrol).to_excel("./result/to_unenrol.xlsx", index=False)
    print("💾 Save to_enrol.xlsx and to_unenrol.xlsx")

    # =========================================================
    # Step 5. Execute enrolment changes
    # =========================================================
    def get_user_id(email, cache):
        if email in cache:
            return cache[email]
        try:
            result = moodle_api(
                "core_user_get_users_by_field", {"field": "email", "values[0]": email}
            )
            if isinstance(result, list) and result:
                userid = result[0]["id"]
                cache[email] = userid
                return userid
            else:
                print(f"Don't find user: {email}")
                return None
        except Exception as e:
            print(f"❌ Get user {email} failed: {e}")
            return None

    user_cache = {}
    STUDENT_ROLE_ID = 5  # Student role is 5

    # ----------------------------
    # Enrol
    # ----------------------------
    # print("\n🚀 开始执行 enrol...")
    # for item in to_enrol:
    #     email = item["email"]
    #     course_id = item["course_id"]
    #     userid = get_user_id(email, user_cache)
    #     if not userid:
    #         continue

    #     enrol_data = {
    #         "enrolments[0][roleid]": STUDENT_ROLE_ID,
    #         "enrolments[0][userid]": userid,
    #         "enrolments[0][courseid]": course_id,
    #     }

    #     try:
    #         moodle_api("enrol_manual_enrol_users", enrol_data)
    #         print(f"✅ Enrolled: {email} -> {item['shortname']}")
    #     except Exception as e:
    #         print(f"❌ Enrol Failed: {email} ({e})")
    #     time.sleep(0.2)

    # ----------------------------
    # Unenrol
    # ----------------------------
    # print("\n🚀 开始执行 unenrol...")
    # for item in to_unenrol:
    #     email = item["email"]
    #     course_id = item["course_id"]
    #     userid = get_user_id(email, user_cache)
    #     if not userid:
    #         continue

    #     unenrol_data = {
    #         "enrolments[0][userid]": userid,
    #         "enrolments[0][courseid]": course_id,
    #     }

    #     try:
    #         moodle_api("enrol_manual_unenrol_users", unenrol_data)
    #         print(f"✅ Unenrolled: {email} -> {item['shortname']}")
    #     except Exception as e:
    #         print(f"❌ Unenrol Failed: {email} ({e})")
    #     time.sleep(0.2)

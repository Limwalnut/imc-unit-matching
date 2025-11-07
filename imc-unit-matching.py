import pandas as pd
import re
from tqdm import tqdm
from dotenv import load_dotenv
import os
import requests
import time
import json

# Load Environment Variables
load_dotenv()

MOODLE_URL = os.getenv("MOODLE_URL")
MOODLE_TOKEN = os.getenv("MOODLE_TOKEN")

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
                    }
                    result.append(course)
    return result


def moodle_api(wsfunction, params):
    """统一封装 Moodle API 调用"""
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


if __name__ == "__main__":
    df_current = pd.read_excel(file_path_current_enrolled_modules)
    df_unit = pd.read_excel(file_path_unit_creation)

    campus_tree = build_campus_tree(df_current)
    module_dict = build_module_dict(df_unit)

    result = generate_mapping(campus_tree, module_dict)

result_df = pd.DataFrame(result)
result_df = result_df.rename(
    columns={"timetable_id": "TimetableID", "short_name": "short_name"}
)

# merge with current enrolled modules
merged_df = pd.merge(df_current, result_df, on="TimetableID", how="inner")

final_df = merged_df[["Email2", "short_name"]].copy()
final_df.rename(columns={"Email2": "email"}, inplace=True)

# ======================================
# Step 2. 通过 API 批量获取 course id
# ======================================
course_map = {}  # shortname -> id
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
            print(f"⚠️ 未找到课程: {shortname}")
        time.sleep(0.2)  # 稍微延时，防止请求太快被限流
    except Exception as e:
        print(f"❌ 获取 {shortname} 失败: {e}")
        continue

print(f"✅ 成功获取 {len(course_map)} 个课程 ID")

# 只保留存在于 Unit Creation 中的课程
final_df = final_df[final_df["short_name"].isin(course_map.keys())]

# 替换 course_id
final_df["course_id"] = final_df["short_name"].map(course_map)

# ======================================
# Step 4. 批量获取每门课程的已 enrol 学生
# Exception - Class "local_o365\webservices\external_format_value" not found
# ======================================
enrolled_data = {}  # { course_id: {email1, email2, ...} }

for short_name, course_id in tqdm(course_map.items(), desc="Fetching enrolled users"):
    try:
        users = moodle_api("core_enrol_get_enrolled_users", {"courseid": course_id})
        print(users)
        enrolled_data[course_id] = {u["email"] for u in users if "email" in u}
    except Exception as e:
        print(f"❌ 获取课程 {short_name} 学生失败: {e}")
        continue

# ======================================
# Step 5. 构建目标 enrol 映射
# ======================================
target_enrol = {}
for _, row in final_df.iterrows():
    email = row["email"]
    course_id = row["course_id"]
    target_enrol.setdefault(course_id, set()).add(email)

# ======================================
# Step 6. 对比当前与目标状态
# ======================================
to_enrol = []
to_unenrol = []

for course_id, shortname in course_map.items():
    current = enrolled_data.get(course_id, set())
    target = target_enrol.get(course_id, set())

    new_users = target - current
    wrong_users = current - target

    for email in new_users:
        to_enrol.append({"email": email, "course_id": course_id, "shortname": shortname})
    for email in wrong_users:
        to_unenrol.append({"email": email, "course_id": course_id, "shortname": shortname})

print(f"✅ 待 Enrol: {len(to_enrol)} 条记录")
print(f"✅ 待 Unenrol: {len(to_unenrol)} 条记录")

# ======================================
# Step 7. 保存结果到 Excel
# ======================================
pd.DataFrame(to_enrol).to_excel("./result/to_enrol.xlsx", index=False)
pd.DataFrame(to_unenrol).to_excel("./result/to_unenrol.xlsx", index=False)

print("💾 已生成 to_enrol.xlsx 与 to_unenrol.xlsx")
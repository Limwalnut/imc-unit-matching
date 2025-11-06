from dotenv import load_dotenv
import os
import requests
import sys

# Load Environment Variables
load_dotenv()

MOODLE_URL = os.getenv("MOODLE_URL")
MOODLE_TOKEN = os.getenv("MOODLE_TOKEN")


# Get Users by Email list
def get_users_by_emails(emails: list):
    """Fetch multiple users' info by email list"""
    params = {
        "wstoken": MOODLE_TOKEN,
        "wsfunction": "core_user_get_users_by_field",
        "moodlewsrestformat": "json",
        "field": "email",
        "values[]": emails,  # 注意这里是 values[] 而不是 values
    }

    try:
        response = requests.get(MOODLE_URL, params=params, timeout=100)
        response.raise_for_status()
        data = response.json()

        if isinstance(data, list) and data:
            filtered = [{"id": user["id"], "email": user["email"]} for user in data]
            print(f"✅ Found {len(filtered)} users.")
            print(filtered)
            return filtered  # 返回完整用户列表
        else:
            print("⚠️ No users found or unexpected response:", data)
            return []
    except requests.exceptions.RequestException as e:
        print("❌ Request failed:", e)
        return []


# Get User Courses
def get_user_courses(user_id: int):
    """Fetch all courses that a user is enrolled in"""
    params = {
        "wstoken": MOODLE_TOKEN,
        "wsfunction": "core_enrol_get_users_courses",
        "moodlewsrestformat": "json",
        "userid": user_id,
    }

    try:
        response = requests.get(MOODLE_URL, params=params, timeout=10)
        response.raise_for_status()
        data = response.json()
        if isinstance(data, list):
            return data
        else:
            print("⚠️ Unexpected response:", data)
            return []
    except requests.exceptions.RequestException as e:
        print("❌ Request failed:", e)
        return []


# Get course_id by course shortname
def get_courseid_by_shortname(shortname: str):
    """Fetch course ID by course shortname"""
    params = {
        "wstoken": MOODLE_TOKEN,
        "wsfunction": "core_course_get_courses_by_field",
        "moodlewsrestformat": "json",
        "field": "shortname",
        "value": shortname,
    }
    response = requests.get(MOODLE_URL, params=params)
    data = response.json()
    if "courses" in data and len(data["courses"]) > 0:
        return data["courses"][0]["id"]
    else:
        print(f"❌ Course '{shortname}' not found.")
        return None


# Main Process
if __name__ == "__main__":
    get_users_by_emails(
        [
            "2025003871@student.imc.edu.au",
            "2023014252@student.imc.edu.au",
            "2024023632@student.imc.edu.au",
        ]
    )
    # get_courseid_by_shortname("2025 Term2 TMGT601")

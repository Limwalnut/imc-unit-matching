from dotenv import load_dotenv
import os
import requests
import sys

# Load Environment Variables
load_dotenv()

MOODLE_URL = os.getenv("MOODLE_URL")
MOODLE_TOKEN = os.getenv("MOODLE_TOKEN")

# Get User by Email
def get_user_by_email(email: str):
    """Fetch user info by email"""
    params = {
        "wstoken": MOODLE_TOKEN,
        "wsfunction": "core_user_get_users_by_field",
        "moodlewsrestformat": "json",
        "field": "email",
        "values[0]": email,
    }

    try:
        response = requests.get(MOODLE_URL, params=params, timeout=10)
        response.raise_for_status()
        data = response.json()
        if isinstance(data, list) and data:
            return data[0]
        else:
            print("⚠️ No user found or unexpected response format:", data)
            return None
    except requests.exceptions.RequestException as e:
        print("❌ Request failed:", e)
        return None

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


# Main Process
if __name__ == "__main__":
    # Change the email to test with different users
    email = "@student.imc.edu.au"

    print(f"🔍 Searching for user: {email}")
    user = get_user_by_email(email)

    if not user:
        sys.exit(0)

    user_id = user.get("id")
    fullname = f"{user.get('firstname', '')} {user.get('lastname', '')}"
    print(f"✅ Found User: {fullname} (ID: {user_id})")
    courses = get_user_courses(user_id)

    if courses:
        print(f"✅ {len(courses)} courses found:")
        for c in courses:
            print(f"   - {c['shortname']} | {c['fullname']}")
    else:
        print("⚠️ No courses found for this user.")
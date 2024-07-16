from datetime import datetime
import os
from supabase import create_client, Client
import win32gui
from dotenv import load_dotenv
import win32process
import psutil
import time
import getpass
import csv
import cv2
import threading
import random
import joblib
import pandas as pd
load_dotenv()


url = os.environ.get("SUPABASE_URL")
key = os.environ.get("SUPABASE_KEY")

last_active_application=None


supabase = create_client(url, key)
role='Developer'


role_app_access = {
    "Developer": ["pycharm64.exe", "python.exe", "code.exe", "sublime_text.exe", "idea64.exe", "eclipse.exe","ApplicationFrameHost.exe"],
    "Designer": ["photoshop.exe", "illustrator.exe", "sketch.exe", "figma.exe"],
    "Finance": ["excel.exe", "sheets.exe", "quickbooks.exe"],
    "Marketing": ["photoshop.exe", "illustrator.exe", "figma.exe", "chrome.exe","explorer.exe"],
}

app_access_limit = {
    "pycharm64.exe": 20,
    "code.exe": 20,
    "sublime_text.exe": 10,
    "idea64.exe": 10,
    "ApplicationFrameHost.exe":5,
    "eclipse.exe": 10,
    "settings.exe": 5,
    "SearchHost.exe": 3,
    "RuntimeBroker.exe": 3,
    "MicrosoftEdge.exe": 5,
    "chrome.exe": 5,
    "Notepad.exe":2,
}


pipeline=joblib.load('pipeline.pkl')

violations = {}
user_name = getpass.getuser()

def capture_photo():

    cap = cv2.VideoCapture(0)

    if not cap.isOpened():
        print("Error: Could not open webcam.")
        return

    try:
        while True:

            ret, frame = cap.read()
            cv2.imshow("Press 's' to save or 'q' to quit", frame)

            key = cv2.waitKey(1) & 0xFF
            if key == ord('s'):  # Press 's' to save the photo
                timestamp = time.strftime("%Y%m%d_%H%M%S")
                filename = f"captured_photo_{timestamp}.jpg"
                cv2.imwrite(filename, frame)
                print(f"Photo saved as {filename}")

            elif key == ord('q'):
                break

            time.sleep(1)

    finally:
        cap.release()
        cv2.destroyAllWindows()
        bucket_name: str = "Samruddha"
        new_file = filename
        r_num = random.randint(1, 10)
        fileimg_name = "user1" + str(r_num) + ".png"
        data = supabase.storage.from_(bucket_name).upload(fileimg_name, new_file)
        print(data)

def get_active_window():
    try:
        hwnd = win32gui.GetForegroundWindow()
        pid = win32process.GetWindowThreadProcessId(hwnd)
        process = psutil.Process(pid[-1])
        app_name = process.name()
        if app_name.lower() in ['searchhost.exe']:
            return None
        if app_name.lower() in ['explorer.exe', 'applicationframehost.exe']:
            window_text = win32gui.GetWindowText(hwnd)
            return window_text
        return app_name
    except Exception as e:
        return None
    
def trigger_critical_alert(user_input):
    current_time = datetime.now().strftime("%Y-%m-%d %H:%M")  
    supabase.table("Alerts").insert(
        {"alert": f'The application {user_input["application"]} has been accessed {user_input["access_count"]} times by {user_name}', "user_name":user_name}
    ).execute()
    capture_photo()
    print(f"Alert: The application {user_input["application"]} has been accessed {user_input["access_count"]} times by {user_name} at {current_time}!")


def log_violations(user_input):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M")
    supabase.table("violations").insert(
        {"user_name": user_input['user_name'], "application": user_input["application"], "violation_count": user_input["access_count"], "timestamp": timestamp, "role": role, "duration": user_input["duration"]}
    ).execute()
    return


def log_application_usage():
    global last_active_application
    app_access_count = {}
    start_time = None
    no_action_start_time = None
    user_name = getpass.getuser()

    while True:
        current_app = get_active_window()
        current_time = time.time()
        if current_app:
            if current_app != last_active_application:
                last_active_application = current_app
                if last_active_application is not None and start_time is not None:
                    duration_minutes = (current_time - start_time) / 60
                    duration_minutes = round(duration_minutes, 2)
                else:
                    duration_minutes = 0  # Initial application or very short duration

                start_time = current_time
                timestamp = datetime.now().strftime("%Y-%m-%d %H:%M")  # Format timestamp

                if current_app in app_access_count:
                    app_access_count[current_app] += 1
                else:
                    app_access_count[current_app] = 1

                print(f"Debug: {current_app} - Access Count: {app_access_count[current_app]}, Access Limit: {app_access_limit.get(current_app, 'No limit set')}")
                
                user_role = supabase.table("user_analysis").select("role").filter("user_name", "eq", user_name).execute().data[0]["role"]

                three_hours_ago = time.time() - 3 * 60 * 60
                if start_time >= three_hours_ago:

                    supabase.table("user_analysis").insert(
                        {"log_time": timestamp, "application": current_app, "user_name": user_name, "duration": duration_minutes, "access_count": app_access_count[current_app],"role":role}
                    ).execute()
                    print(timestamp, current_app, user_name, f"{duration_minutes} minutes", f"Access Count: {app_access_count[current_app]}")
                no_action_start_time = None
                return {"id":[1],"log_time": [timestamp], "application": [current_app], "user_name": [user_name], "duration": [duration_minutes], "access_count": [app_access_count[current_app]], "role": [role]}
        else:

            if no_action_start_time is None:
                no_action_start_time = current_time
            else:

                if current_time - no_action_start_time >= 60:
                    print("Alert: No process has been in action for 1 minute!")
        time.sleep(1) 

def process_log_data():
  global last_active_application
  while True:  
    user_input = log_application_usage()  
    input_df = pd.DataFrame(user_input)
    print(input_df)
    input_df['log_time'] = pd.to_datetime(input_df['log_time'])
    input_df['hour'] = input_df['log_time'].dt.hour
    input_df['day_of_week'] = input_df['log_time'].dt.dayofweek
    input_df.drop(['log_time', 'id', 'user_name'], axis=1, inplace=True)
    
    print(input_df.shape)
    prediction = pipeline.predict(input_df)
    print(f'Predicted Label: {prediction[0]}')
    if prediction[0] == 1:
        log_violations(user_input)  
    elif prediction[0] == 2:
        trigger_critical_alert(user_input)  

if __name__ == "__main__":
    process_log_data()
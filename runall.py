import schedule
import time
import subprocess

def run_scripts():
    scripts = ["dailyPMPlan.py", "ScheqReports_copy.py", "combo.py"]  # Replace with your script names
    for script in scripts:
        try:
            subprocess.run(["python", script], check=True)
            print(f"Successfully ran {script}")
        except subprocess.CalledProcessError as e:
            print(f"Error running {script}: {e}")

# Define the time to run the scripts (24-hour format, e.g., '14:30' for 2:30 PM)
schedule_time = "13:12"  # Change this to your desired time
schedule.every().day.at(schedule_time).do(run_scripts)

print(f"Scheduled scripts to run daily at {schedule_time}")

while True:
    schedule.run_pending()
    time.sleep(5)  # Check every minute

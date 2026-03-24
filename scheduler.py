import schedule
import time
import os

def run_app():
    os.system("streamlit run app.py")

# ⏰ Set time (24-hour format)
schedule.every().day.at("10:00").do(run_app)

while True:
    schedule.run_pending()
    time.sleep(60)
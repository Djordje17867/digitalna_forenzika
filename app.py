import psutil
import time
from datetime import datetime
import platform
import smtplib
from fpdf import FPDF
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from pynput import mouse, keyboard  # For tracking clicks and keystrokes
from dotenv import load_dotenv
import os
import subprocess

load_dotenv()

EMAIL_HOST = os.getenv("EMAIL_HOST")
EMAIL_PORT = os.getenv("EMAIL_PORT")
EMAIL_USERNAME = os.getenv("EMAIL_USERNAME")
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD")
EMAIL_RECEIVER = os.getenv("EMAIL_RECEIVER")

# Dictionary to track mouse and keyboard activity per hour
activity_per_hour = {hour: (0, 0) for hour in range(9, 22)}
usage_log = {}

def is_within_work_hours():
    current_time = datetime.now().time()
    start_time = datetime.strptime("09:00", "%H:%M").time()
    end_time = datetime.strptime("22:00", "%H:%M").time()
    return start_time <= current_time <= end_time

def get_active_window():
    current_os = platform.system()

    if current_os == "Darwin":
        script = '''
        tell application "System Events"
            set frontApp to first application process whose frontmost is true
            set appName to name of frontApp
            set windowTitle to ""
            if appName is "Electron" then
                try
                    set windowTitle to name of first window of frontApp
                end try
            end if
            return {appName, windowTitle}
        end tell
        '''
        try:
            result = subprocess.run(['osascript', '-e', script], capture_output=True, text=True)
            output = result.stdout.strip()
            if output:
                values = output.split(", ")
                app_name = values[0].replace(",", "")
                if app_name.__contains__("Electron"):
                    return "Visual Studio Code"
                return app_name
            else:
                return "No active application found"
        except subprocess.CalledProcessError as e:
            return f"Error: {e}"

    elif current_os == "Windows":
        import pygetwindow as gw
        window = gw.getActiveWindow()
        if window is not None:
            return window.title
        return None

    else:
        return "Unsupported OS"

def get_current_hour():
    return datetime.now().hour

def log_application_usage():
    print("""Track app usage between work hours.""")
    while is_within_work_hours():
        active_window = get_active_window()
        if active_window:
            if active_window not in usage_log:
                usage_log[active_window] = {'start_time': datetime.now(), 'total_time': 0}
            usage_log[active_window]['total_time'] += 10  # Increment time by 10 seconds

        time.sleep(10)  # Check every 10 seconds

    generate_pdf_report()  # At 5 PM, generate PDF and stop

def on_click(x, y, button, pressed):
    if pressed:
        current_hour = get_current_hour()
        mouse_clicks, keyboard_presses = activity_per_hour[current_hour]
        activity_per_hour[current_hour] = (mouse_clicks + 1, keyboard_presses)

def on_press(key):
    current_hour = get_current_hour()
    mouse_clicks, keyboard_presses = activity_per_hour[current_hour]
    activity_per_hour[current_hour] = (mouse_clicks, keyboard_presses + 1)

def generate_pdf_report():
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()

    pdf.set_font("Arial", size=12)
    pdf.cell(200, 10, txt="Application Usage Report", ln=True, align='C')
    pdf.ln(10)

    # Write the log details to the PDF
    for app, info in usage_log.items():
        total_seconds = info['total_time']
        total_minutes = total_seconds // 60 
        hours = total_minutes // 60
        minutes = total_minutes % 60 

        if hours > 0:
          output = f"App: {app}, Total Time Used: {hours} hours and {minutes} minutes"
        else:
         output = f"App: {app}, Total Time Used: {minutes} minutes"

    pdf.cell(200, 10, txt=output, ln=True)

    pdf.ln(10)
    pdf.cell(200, 10, txt="Mouse Clicks and Keyboard Presses by Hour:", ln=True)
    for hour, (mouse_clicks, keyboard_presses) in activity_per_hour.items():
        if mouse_clicks > 0 or keyboard_presses > 0:
            pdf.cell(200, 10, txt=f"{hour}:00 - {hour + 1}:00 -> Mouse Clicks: {mouse_clicks}, Keyboard Presses: {keyboard_presses}", ln=True)

    timestamp = datetime.now().strftime("%Y-%m-%d")
    pdf_output = f"app_usage_report_{timestamp}.pdf"
    pdf_file_path = os.path.join('Reports', pdf_output)
    pdf.output(pdf_file_path)
    print(f"PDF report saved as {pdf_output}")

    send_email(pdf_file_path)

def send_email(pdf_filename):
    print("""Sending the PDF report via email.""")
    msg = MIMEMultipart()
    msg['From'] = EMAIL_USERNAME
    msg['To'] = EMAIL_RECEIVER
    msg['Subject'] = "Daily Application Usage Report"

    # Attach the PDF file
    with open(pdf_filename, 'rb') as attachment:
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename={pdf_filename}')
        msg.attach(part)

    # Sending the email via SMTP
    try:
        with smtplib.SMTP(EMAIL_HOST, EMAIL_PORT) as server:
            server.starttls()
            server.login(EMAIL_USERNAME, EMAIL_PASSWORD)
            server.sendmail(EMAIL_USERNAME, EMAIL_RECEIVER, msg.as_string())
            print(f"Email sent to {EMAIL_RECEIVER}")
    except Exception as e:
        print(f"Failed to send email: {e}")

if __name__ == "__main__":
    # Set up listeners for mouse and keyboard tracking
    mouse_listener = mouse.Listener(on_click=on_click)
    keyboard_listener = keyboard.Listener(on_press=on_press)
    
    mouse_listener.start()
    keyboard_listener.start()
    
    try:
        while is_within_work_hours():
            log_application_usage()
    except KeyboardInterrupt:
        # Output the log when stopped (manually stop before 5 PM)
        generate_pdf_report()
    
    # Stop listeners when work hours end or program ends
    mouse_listener.stop()
    keyboard_listener.stop()

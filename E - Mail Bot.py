import os
import re
import random
import dns.resolver 
import pandas as pd
import smtplib
from sys import exit
from config import *
from email.message import EmailMessage 
from datetime import datetime, timedelta
EMAIL = E_MAIL                      # MY GMAIL ADDRESS
PASSWORD = PASSWORD            # MY 16-CHARACTER APP PASSWORD
SMTP_SERVER = "smtp.gmail.com"  
SMTP_PORT = 465

try:
    output_folder = os.path.join(os.environ["USERPROFILE"], "Downloads")
    EXCEL_FILE = os.path.join(output_folder, "Data Sheet.xlsx")

    if not os.path.exists(EXCEL_FILE):
        print(f"ERROR: File not found at → {EXCEL_FILE}")
        print(f"ERROR: File not found at → {EXCEL_FILE}")
        exit(1)

except KeyError:
    print("ERROR: USERPROFILE environment variable not found (are you on Windows?).")
    exit(1)

except Exception as e:
    print(f"ERROR: Unexpected issue while setting file path:\n{e}")
    exit(1)

festivals_2025 = {
    "02-10": "Gandhi Jayanti", 
    "05-09": "Teachers Day",
    "22-09": "Navratri Begins",
    "02-10": "Mahatma Gandhi Jayanti",
    "02-10": "Dussehra",
    "20-10": "Diwali",
    "23-10": "Bhai Duj",
    "11-11": "Chhath Puja",
    "25-12": "Christmas",
}

# My Firewall For E-mail Address Checking
def check_email(email):
    email = email.strip()
    pattern = r'^[\w\.-]+@[\w\.-]+\.\w+$'
    
    if not re.match(pattern, email):
        print(f"\nInvalid email format: {email}")
        return False
    
    domain = email.split('@')[-1].strip()
    
    try:
        mx_records = dns.resolver.resolve(domain, 'MX')
        #print(f"\nDomain verified: {email} → {mx_records[0].exchange}")
        return True
    
    except dns.resolver.NXDOMAIN:
        print(f"\nInvalid domain for {email}: Domain does not exist.")
    
    except dns.resolver.NoAnswer:
        print(f"\nNo MX record found for {email}.")
    
    except dns.resolver.NoNameservers:
        print(f"\nNo name servers for domain of {email}.")
    
    except Exception as e:
        print(f"\nUnknown error for {email}: {e}")
    
    return False

# Load HTML templates safely
def load_template(filename):
    try:
        path = os.path.join(os.path.dirname(__file__),"templet", filename)
        with open(path, 'r', encoding='utf-8') as file:
            return file.read()
    
    except Exception as e:
        print(f"Error loading HTML template {filename}: {e}")
        return None

TEMPLATE_JOIN = load_template("Thank_You_Mail.html")
TEMPLATE_BDAY = load_template("Birthday_Wishing_Mail.html")
TEMPLATE_FESTIVAL = load_template("Festival_Mail.html")
TWO_FACTOR_AUTHENTICATION_PROCESS = load_template("Two_Factor_Authentication.html")

# Email sended Record
def records_of_email_sended(From,To,Subject):
    try:
        path = os.path.join(os.path.dirname(__file__), "Trec_Records_Of_Email_Sended.txt")
        with open(path, "a", encoding="utf-8") as file:
            now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            file.write(f"\n\n\nDate : [{now}]\n\n")
            file.write(f"--> Sender Email Address: {From}\n")
            file.write(f"--> Receiver Email Address: {To}\n")
            file.write(f"--> Subject: {Subject}\n")
    except Exception as e:
        print(f"Error writing log: {e}")

def send_email_html(to, subject, html_content):
    try:
        msg = EmailMessage()
        msg['From'] = EMAIL
        msg['To'] = to
        msg['Subject'] = subject
        msg['Reply-To'] = RESPONS_MAILE      # MY ALLWAYS REPLY PROTON MAIL ADDRES 
        msg.set_content("This is an HTML email.")
        msg.add_alternative(html_content, subtype='html')

        with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT) as smtp:
            smtp.login(EMAIL, PASSWORD)
            smtp.send_message(msg)
            records_of_email_sended(EMAIL,to,subject)
            print(f"\n--> Email sent to {to} | Subject: {subject}\n")
        return True
    
    except Exception as e:
        print(f"ERROR while sending email to {to} : {e}")
        return False

# Login Attempt Record
def log_login_attempt(status, entered_otp="N/A", correct_otp="N/A"):
    try:
        path = os.path.join(os.path.dirname(__file__), "Trec_Records_Of_Login_Attempt.txt")
        with open(path, "a", encoding="utf-8") as file:
            now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            file.write(f"\n\n\nDate : [{now}]\n\n")
            file.write(f"--> Status: {status}\n")
            file.write(f"--> Entered OTP: {entered_otp}\n")
            file.write(f"--> Correct OTP: {correct_otp}\n")
    except Exception as e:
        print(f"Error writing log: {e}")

# 2 Factor Authentication Process
def two_factor_authentication_process():
    print("\nLogin Required\n1. Yes\n2. No")
    choice = input("Enter choice: ").strip()

    if choice != "1":
        print("Exiting program...")
        log_login_attempt("CANCELLED")
        exit(0)

    # Generate OTP
    otp = random.randint(100000, 999999)
    html_template = TWO_FACTOR_AUTHENTICATION_PROCESS
    html_content = html_template.replace("******", str(otp))

    # Send OTP to your own email
    if send_email_html(RESPONS_MAILE, "Two Factor Authentication Process", html_content):
        print("OTP sent to your email.")

        entered_otp = input("\nEnter the 8-Digit OTP received in your email: ").strip()
        if entered_otp == str(otp):
            print("\nOTP Verified Successfully!\n")
            log_login_attempt("SUCCESS", entered_otp, otp)

            print("Are you logging in?\n1. Yes, I am in\n2. No, I am not in")
            final_choice = input("Enter choice: ").strip()

            if final_choice == "1":
                return True
            else:
                print("Exiting program...")
                log_login_attempt("CANCELLED_AFTER_VERIFY", entered_otp,otp)
                exit(0)
        else:
            print("Invalid OTP! Exiting...")
            log_login_attempt("FAIL", entered_otp,otp)
            exit(0)
    else:
        print("Failed to send OTP. Exiting...")
        log_login_attempt("OTP_SEND_FAIL")
        exit(1)

def main():
    try:
        now = datetime.now()
        today = now.strftime("%d-%m")
        try:
            df = pd.read_excel(EXCEL_FILE, parse_dates=['Birthday'])

        except Exception as e:
            print(f"ERROR reading EXCEL file : {e}")
            exit(1)

        birthday_sent = False
        joining_sent = False

        if today in festivals_2025:
            print(f"\nToday is {festivals_2025[today]}!\n")
        
        else:
            print("\nNo festival today.\n")

        # Joining Emaile (Last 24 hrs)
        now = datetime.now().replace(microsecond=0)
        print("\nChecking for new joiners...\n")
        for _, row in df.iterrows():
            try:
                if pd.notnull(row['Timestamp']) and TEMPLATE_JOIN:
                    join_time = pd.to_datetime(row['Timestamp'], format="%m/%d/%Y %H:%M:%S", errors='coerce')
                    if pd.notnull(join_time):
                        if timedelta(hours=0) <= (now - join_time) <= timedelta(hours=24):
                            name = str(row['Name'])
                            email = str(row['Email'])
                            print(f"--> New Joiner: {name} | Joined: {join_time}")
                            join_str = join_time.strftime("%d-%m-%Y %H:%M:%S")
                            html = TEMPLATE_JOIN.replace("{name}", name).replace("{join_time_str}", join_str)
                            if check_email(email):
                                if send_email_html(email, f"Thank You for Joining Us, {name}", html):
                                    joining_sent = True
            
            except Exception as e:
                print(f"Join Email Error for {row.get('Name','Unknown')}: {e}")

       # Check for birthdays
        print("\nChecking for birthdays...\n")
        for _, row in df.iterrows():
            try:
                name = str(row['Name'])
                email = str(row['Email'])
                if pd.notnull(row['Birthday']) and TEMPLATE_BDAY:
                    bday = row['Birthday'].strftime("%d-%m")
                    if bday == today:
                        print(f"\n--> Happy B'day to {name}!")
                        html = TEMPLATE_BDAY.replace("{name}", name).replace("{festival}", "Your Birthday")
                        if check_email(email):
                            if send_email_html(email, f"Happy Birthday, {name}!", html):
                                birthday_sent = True
            
            except Exception as e:
                print(f"Birthday Check Error for {row.get('Name','Unknown')}: {e}")

        # Check for today's festival
        if today in festivals_2025 and TEMPLATE_FESTIVAL:
            print("\nSending Festival Mails...")
            for _, row in df.iterrows():
                try:
                    festival = festivals_2025[today]
                    name = str(row['Name'])
                    email = str(row['Email'])
                    html = TEMPLATE_FESTIVAL.replace("{name}", name).replace("{festival}", festival)
                    if check_email(email):
                        send_email_html(email, f"Warm Wishes on {festival}", html)
                
                except Exception as e:
                    print(f"Festival Email Error for {row.get('Name','Unknown')}: {e}")

        if not birthday_sent and today not in festivals_2025 and not joining_sent:
            print("\nNo birthdays today\nNo festivals today\nNo new joiners in last 24 hrs\n")

    except Exception as e:
        print(f"ERROR in main Execution : {e}")
        exit(1)

    finally:
        print("|--------------------------------------------------------|")
        print("\n|-----Program complete. You may close VS Code now.-------|\n")
        print("|--------------------------------------------------------|")
# Start
if __name__ == "__main__":
    two_factor_authentication_process()
    main()
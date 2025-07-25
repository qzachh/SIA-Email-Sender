#!/usr/bin/env python3
import pandas as pd
import subprocess

def send_email(poc_name, firm_name, recipient_email):
    subject = "Seeking advice on architectural documentation challenges"

    # Adjust possessive form
    firm_name_clean = firm_name.strip()
    firm_possessive = f"{firm_name_clean}'" if firm_name_clean.endswith('s') else f"{firm_name_clean}'s"

    body = f"""Hi {poc_name},

I’m Kai, a student from NUS who recently transitioned from Architecture to Computing. While studying Architecture, I often struggled with the time-consuming documentation tasks that took away from the design process.

As part of my project, I’m learning more about the documentation process and the challenges architects face. I was wondering if you could help me understand {firm_possessive} experience with the following:

1️⃣ Workflow Pain Points
What are the most tedious or frustrating parts of your documentation process today?

2️⃣ Compliance Checking
How do you check your drawings for compliance (against project requirements / regulations / client-specific preferences) — and where do issues tend to slip through?

3️⃣ Generating Submission Drawings
Do you find it repetitive or manual to prepare different versions of drawings for various purposes (e.g. electrical, client presentation, tender, authority submission)?

I’d greatly appreciate any insights you could share. If you’re open to it, I’d love to chat more about your thoughts.

Thank you so much!
Kai
NUS Computing
SG +65 9776 3340 | VN +84 3693 89242"""

    # Escape quotes for AppleScript
    esc_body = body.replace('"', '\\"')
    esc_subject = subject.replace('"', '\\"')

    # AppleScript to send via Outlook
    applescript = f'''
    tell application "Microsoft Outlook"
        set newMessage to make new outgoing message with properties {{subject:"{esc_subject}", content:"{esc_body}"}}
        tell newMessage
            make new to recipient at end of to recipients with properties {{email address:{{address:"{recipient_email}"}}}}
            make new cc recipient at end of cc recipients with properties {{email address:{{address:"joelleo@comp.nus.edu.sg"}}}}
            send
        end tell
    end tell
    '''

    subprocess.run(['osascript', '-e', applescript], check=True)

def main():
    df = pd.read_csv("Auto_Architecture_Firms_1.csv", dtype=str)

    for _, row in df.iterrows():
        poc_name = row.get("POC_Name", "").strip()
        firm_name = row.get("Firm_Name", "").strip()
        recipient_email = row.get("Firm_Email", "").strip()

        if poc_name and firm_name and recipient_email:
            try:
                send_email(poc_name, firm_name, recipient_email)
                print(f"✅ Sent to {poc_name} at {recipient_email}")
            except Exception as e:
                print(f"❌ Failed to send to {recipient_email}: {e}")

if __name__ == "__main__":
    main()

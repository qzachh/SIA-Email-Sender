#!/usr/bin/env python3
import pandas as pd
import subprocess

def send_email(email, salutation, first_name, company, projects):
    subject = "Trying to help architects spend less time on documentation — would love your input"

    # Choose template based on Projects value
    if projects.strip() == "-":
        body = f"""Hi {salutation} {first_name},

I'm Kai. I switched from architecture to computing because endless paperwork & compliances were crowding out design time. From talking with other architects, I know teams at firms like {company} can easily get buried in drawings, specs, and approvals.

I’m prototyping a workflow tool for architects. Would you be open to share which parts of your documentation process feels most painful right now?

Happy to chat by email, call, or whatever works for you. Thank you!

Appreciate your insight,
Kai
NUS Computing
SG +65 9776 3340 | VN +84 3693 89242"""
    else:
        body = f"""Hi {salutation} {first_name},

I'm Kai. I switched from architecture to computing because endless paperwork & compliances were crowding out design time. {company}'s recent project, {projects}, reminded me how things can become tedious when documents pile up.

I’m prototyping a workflow tool for architects. Could you share which parts of your documentation process feel most painful right now?

Happy to chat by email, call, or whatever works for you.

Appreciate your insight,
Kai
NUS Computing
SG +65 9776 3340 | VN +84 3693 89242"""

    # Escape for AppleScript
    esc_body = body.replace('"', '\\"')
    applescript = f'''
tell application "Microsoft Outlook"
    set msg to make new outgoing message with properties {{subject:"{subject}", content:"{esc_body}"}}
    tell msg
        make new to recipient at end of to recipients with properties {{email address:{{address:"{email}"}}}}
        make new cc recipient at end of cc recipients with properties {{email address:{{address:"joelleo@comp.nus.edu.sg"}}}}
        send
    end tell
end tell
'''
    subprocess.run(['osascript', '-e', applescript], check=True)

def main():
    df = pd.read_csv("apollo_all.csv", dtype=str)
    for _, row in df.iterrows():
        email      = row.get("Email", "").strip()
        salutation = row.get("Mr/Ms", "").strip()
        first_name = row.get("First Name", "").strip()
        company    = row.get("Company", "").strip()
        projects   = row.get("Projects", "").strip()
        if email:
            send_email(email, salutation, first_name, company, projects)

if __name__ == "__main__":
    main()

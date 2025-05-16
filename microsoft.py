import pandas as pd
import smtplib
import time
import email.utils
import threading
import random
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# CONFIGURATION
XLSX_PATH      = "test Email.xlsx"
INITIAL_GAP    = 60           # 1 minute between initial sends
FOLLOWUP_DELAY = 86400        # seconds (1 day)
SENDER = {
    "email":      "neal@filldesignprojects.website",
    "password":   "Fdg@9874#",
    "smtp_host":  "smtp.office365.com",
    "smtp_port":  587
}

def spin_email_template(person_name, company, is_followup=False, followup_number=None):
    greetings = [f"Hi {person_name},", f"Hello {person_name},", f"Dear {person_name},"]
    s1 = random.choice([
        "I see you booked your new domain, marking an important step toward establishing a strong online presence.",
        "I noticed you secured your new domain—an essential move toward building a reliable online identity.",
        "I noticed you secured your domain. This marks the beginning of your online journey."
    ])
    s2 = random.choice([
        "In the past six months, we’ve worked with several businesses to build websites, improve their search performance, and refine their social media presence. Consider how a well-designed digital platform can support your goals.",
        "Over the past six months, we’ve assisted a number of companies with website design, search optimization, and social media strategy. Think about how a customized digital solution could benefit your business.",
        "Recently, we’ve helped several businesses develop websites, enhance their search performance, and improve their social media efforts. Imagine a digital solution that aligns with your business needs."
    ])
    s3 = random.choice([
        "I’m contacting you personally to share how our services may be of benefit. Please take a moment to watch the brief video I recorded, which explains our approach.",
        "I’m contacting you directly to share more about our services. I’ve prepared a brief video introduction outlining our process.",
        "I’m reaching out personally to share how our services may help. I’ve recorded a short video to introduce myself and explain our approach."
    ])
    greeting = random.choice(greetings)
    extra = (f"\nThis is follow-up #{followup_number}. Just checking in regarding my previous email."
             if is_followup and followup_number else "")
    loom_link = "https://www.loom.com/share/35049856e0e447e8ada77a44a1297342"

    text = f"""{greeting}

{s1}

{s2}

{s3}

{extra}

{loom_link}

Looking forward to hearing from you.

Best regards,
Neal
https://filldesigngroup.com/
"""
    html = f"""\
<html><body>
  <p>{greeting}</p>
  <p>{s1}</p><p>{s2}</p><p>{s3}</p>
  {f"<p>{extra}</p>" if extra else ""}
  <div><a href="{loom_link}">
    <img style="max-width:300px;"
         src="https://cdn.loom.com/sessions/thumbnails/35049856e0e447e8ada77a44a1297342-b9abd9c74a5b4e39-full-play.gif"
         alt="Watch Video"></a></div>
  <p>Looking forward to hearing from you.<br>
     Best regards,<br>
     Neal<br>
     <a href="https://filldesigngroup.com/">Fill Design Group</a>
  </p>
</body></html>
"""
    return text, html

def choose_subject(company):
    return random.choice([
        "Question for {Company}",
        "See this for {Company}",
        "Quick Question for {Company}"
    ]).format(Company=company)

def send_mail(to_addr, subject, text, html, orig_msg_id=None, is_followup=False):
    msg = MIMEMultipart("alternative")
    msg["From"] = SENDER["email"]
    msg["To"]   = to_addr
    msg["Subject"] = ("Re: " if is_followup else "") + subject
    if is_followup and orig_msg_id:
        msg["In-Reply-To"] = orig_msg_id
        msg["References"]  = orig_msg_id

    msg_id = email.utils.make_msgid()
    if not is_followup:
        msg["Message-ID"] = msg_id

    msg.attach(MIMEText(text, "plain"))
    msg.attach(MIMEText(html, "html"))

    with smtplib.SMTP(SENDER["smtp_host"], SENDER["smtp_port"]) as server:
        server.ehlo()
        server.starttls()
        server.login(SENDER["email"], SENDER["password"])
        server.sendmail(SENDER["email"], to_addr, msg.as_string())

    print(f"{'Follow-up' if is_followup else 'Initial'} sent to {to_addr}")
    return msg_id, subject

def schedule_followups(to_addr, name, company, orig_msg_id, orig_subject):
    # Wait 1 day, then Follow‑up #1
    time.sleep(FOLLOWUP_DELAY)
    t1_text, t1_html = spin_email_template(name, company, True, 1)
    send_mail(to_addr, orig_subject, t1_text, t1_html, orig_msg_id, True)

    # Wait another day, then Follow‑up #2
    time.sleep(FOLLOWUP_DELAY)
    t2_text, t2_html = spin_email_template(name, company, True, 2)
    send_mail(to_addr, orig_subject, t2_text, t2_html, orig_msg_id, True)

def main():
    df = pd.read_excel(XLSX_PATH, engine="openpyxl")
    for _, row in df.iterrows():
        company, name, email_addr = row["company"], row["name"], row["email"]

        # send initial
        text, html = spin_email_template(name, company)
        subject = choose_subject(company)
        msg_id, subj = send_mail(email_addr, subject, text, html)

        # spawn follow-up thread
        threading.Thread(
            target=schedule_followups,
            args=(email_addr, name, company, msg_id, subj),
            daemon=True
        ).start()

        # wait before next initial
        time.sleep(INITIAL_GAP)

if __name__ == "__main__":
    main()

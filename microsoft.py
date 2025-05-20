import smtplib
import time
import random
import email.utils
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# Set TESTING_MODE = True for testing (short delays) and False for production (1 and 2 days)
TESTING_MODE = False
followup_delay = 10 if TESTING_MODE else 86400  # 86400 seconds = 1 day

SENDER_ACCOUNT = {
    "email": "neal@filldesigngroup.net",
    "password": "Fdg@9874#",
    "smtp_server": "smtp.office365.com",
    "smtp_port": 587
}

def get_random_sender():
    return SENDER_ACCOUNT

def spin_email_template(person_name, company, is_followup=False, followup_number=None):
    greetings = [f"Hi {person_name},", f"Hello {person_name},", f"Dear {person_name},"]
    sentence1 = random.choice([
        "I see you booked your new domain, marking an important step toward establishing a strong online presence.",
        "I noticed you secured your new domain—an essential move toward building a reliable online identity.",
        "I noticed you secured your domain. This marks the beginning of your online journey."
    ])
    sentence2 = random.choice([
        "In the past six months, we’ve worked with several businesses to build websites, improve their search performance, and refine their social media presence. Consider how a well-designed digital platform can support your goals.",
        "Over the past six months, we’ve assisted a number of companies with website design, search optimization, and social media strategy. Think about how a customized digital solution could benefit your business.",
        "Recently, we’ve helped several businesses develop websites, enhance their search performance, and improve their social media efforts. Imagine a digital solution that aligns with your business needs."
    ])
    sentence3 = random.choice([
        "I’m contacting you personally to share how our services may be of benefit. Please take a moment to watch the brief video I recorded, which explains our approach.",
        "I’m contacting you directly to share more about our services. I’ve prepared a brief video introduction outlining our process.",
        "I’m reaching out personally to share how our services may help. I’ve recorded a short video to introduce myself and explain our approach."
    ])
    extra = f"\nThis is follow-up #{followup_number}. Just checking in regarding my previous email." if is_followup and followup_number else ""
    loom_link = "https://www.loom.com/share/1915f664b7f145f193d7b0fd6873ecb1"

    # Plain-text body
    text = f"""{random.choice(greetings)}

{sentence1}

{sentence2}

{sentence3}

{extra}

{loom_link}

Looking forward to hearing from you.

Best regards,
Neal
https://filldesigngroup.com/
"""

    # HTML body (with Loom GIF thumbnail)
    html = f"""\
<html><body>
  <p>{random.choice(greetings)}</p>
  <p>{sentence1}</p>
  <p>{sentence2}</p>
  <p>{sentence3}</p>
  {f"<p>{extra}</p>" if extra else ""}
  <div>
    <a href="{loom_link}">
      <img style="max-width:300px;"
           src="https://cdn.loom.com/sessions/thumbnails/1915f664b7f145f193d7b0fd6873ecb1-12ee91ac978e3ba5-full-play.gif"
           alt="Watch Video">
    </a>
  </div>
  <p>Looking forward to hearing from you.<br>
     Best regards,<br>Neal<br>
     <a href="https://filldesigngroup.com/">Fill Design Group</a></p>
</body></html>
"""
    return text, html

def choose_subject(company):
    return random.choice([
        "Question for {Company}",
        "See this for {Company}",
        "Quick Question for {Company}"
    ]).format(Company=company)

def check_reply(email_address):
    return False  # placeholder; implement your own reply-check logic

def send_initial_email(row):
    company = row['company']
    name    = row['name']
    to_addr = row['email']
    subject = choose_subject(company)
    text, html = spin_email_template(name, company)

    sender = get_random_sender()
    msg = MIMEMultipart('alternative')
    msg['From']     = sender['email']
    msg['To']       = to_addr
    msg['Subject']  = subject
    msg['Reply-To'] = sender['email']
    msg['Date']     = email.utils.formatdate(localtime=True)
    msg['X-Mailer'] = "FillDesignMailer/1.0"

    msg.attach(MIMEText(text, 'plain'))
    msg.attach(MIMEText(html, 'html'))

    try:
        with smtplib.SMTP(sender['smtp_server'], sender['smtp_port'], timeout=10) as server:
            server.ehlo()
            server.starttls()
            server.ehlo()
            server.login(sender['email'], sender['password'])
            server.sendmail(sender['email'], to_addr, msg.as_string())
            print(f"Initial email sent to {to_addr} from {sender['email']}")
    except Exception as e:
        print(f"Error sending to {to_addr}: {e}")
        return None, None, None

    return msg['Message-ID'], subject, sender

def send_followup(to_addr, msg_id, name, company, num, sender, orig_subj):
    text, html = spin_email_template(name, company, True, num)
    msg = MIMEMultipart('alternative')
    msg['From']        = sender['email']
    msg['To']          = to_addr
    msg['Subject']     = "Re: " + orig_subj
    msg['In-Reply-To'] = msg_id
    msg['References']  = msg_id
    msg['Reply-To']    = sender['email']
    msg['Date']        = email.utils.formatdate(localtime=True)
    msg['X-Mailer']    = "FillDesignMailer/1.0"

    msg.attach(MIMEText(text, 'plain'))
    msg.attach(MIMEText(html, 'html'))

    try:
        with smtplib.SMTP(sender['smtp_server'], sender['smtp_port'], timeout=10) as server:
            server.ehlo()
            server.starttls()
            server.ehlo()
            server.login(sender['email'], sender['password'])
            server.sendmail(sender['email'], to_addr, msg.as_string())
            print(f"Follow‑up #{num} sent to {to_addr}")
    except Exception as e:
        print(f"Error sending follow‑up #{num} to {to_addr}: {e}")

def followup_scheduler(to_addr, msg_id, name, company, sender, subj):
    # Wait 1 day, then send follow‑up #1 if no reply
    time.sleep(followup_delay)
    if not check_reply(to_addr):
        send_followup(to_addr, msg_id, name, company, 1, sender, subj)
    else:
        return

    # Wait another day, then send follow‑up #2 if still no reply
    time.sleep(followup_delay)
    if not check_reply(to_addr):
        send_followup(to_addr, msg_id, name, company, 2, sender, subj)

def send_emails(xlsx_path):
    df = pd.read_excel(xlsx_path, engine='openpyxl')
    for _, row in df.iterrows():
        company = row['company']
        name    = row['name']
        email   = row['email']
        print(f"Processing: {company} | {name} | {email}")

        msg_id, subj, sender = send_initial_email(row)
        if not msg_id:
            continue

        # Run follow‑ups in the same thread (blocking)
        if TESTING_MODE:
            # Short delays when testing
            followup_scheduler(email, msg_id, name, company, sender, subj)
        else:
            # 1‑day and 2‑day delays in production
            followup_scheduler(email, msg_id, name, company, sender, subj)

        # Pause 60 seconds before sending to next recipient
        time.sleep(60)

if __name__ == "__main__":
    send_emails("test Email.xlsx")

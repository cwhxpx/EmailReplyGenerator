from openai import OpenAI
import win32com.client
import tkinter as tk
import config
import imapclient
import pyzmail
import email
from email.policy import default
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import smtplib

# Define your Outlook credentials
email_account = config.OUTLOOK_USERNAME
email_pwd = config.OUTLOOK_PASSWORD

openai_client = OpenAI(
    api_key=config.OPENAI_API_KEY
)
def last_10_emails():
    # 连接到IMAP服务器
    server = imapclient.IMAPClient('imap-mail.outlook.com', ssl=True)
    server.login('sun.shuwei@outlook.com', 'fdydzvbumnfrkxco')
    # 选择邮箱文件夹
    server.select_folder('INBOX')
    # 搜索邮件（此处为搜索所有未删除的邮件）
    messages = server.search(['NOT', 'DELETED'])
    # 获取最新的10个邮件的UID
    latest_messages = messages[-10:] if len(messages) > 10 else messages
    # 获取这些邮件的内容和标志
    response = server.fetch(latest_messages, ['BODY[]', 'FLAGS'])
    emails = []
    global email_subjects_dict
    # Dictionary to store email subjects and their corresponding msg_ids
    email_subjects_dict = {}
    # 处理并显示邮件
    for msgid, data in response.items():
        email_message = email.message_from_bytes(data[b'BODY[]'], policy=default)
        subject = email_message['subject']
        from_ = email_message['from']
        to = email_message['to']
        subject_info = f"Subject: {subject}, From: {from_}, To: {to}"
        emails.append(subject_info)
        # Store the subject info and its corresponding msg_id
        email_subjects_dict[subject_info] = msgid
    # 断开连接
    server.logout()
    return emails

root = tk.Tk()
root.title("Outlook Emails")
root.geometry("300x300")
email_subjects = last_10_emails()
for e in email_subjects:
    print(e)
selected_subject = tk.StringVar()
dropdown = tk.OptionMenu(root, selected_subject, *email_subjects)
dropdown.pack()
label = tk.Label(root, text="")
label.pack()

def reply():
    # Connect to Outlook
    # 连接到IMAP服务器
    server = imapclient.IMAPClient('imap-mail.outlook.com', ssl=True)
    server.login(email_account, email_pwd)
    server.select_folder('INBOX')
    # Get the selected email
    selected_subject_value = selected_subject.get()
    print(selected_subject_value)
    msg_id = email_subjects_dict.get(selected_subject_value)
    messages = server.search(['SUBJECT', selected_subject_value])
    if not msg_id:
        print("Email not found")
        return
    msg_data = server.fetch(msg_id, ['RFC822'])
    raw_email = msg_data[msg_id][b'RFC822']
    email_message = email.message_from_bytes(raw_email, policy=default)
    email_body = ""
    if email_message.is_multipart():
        for part in email_message.walk():
            if part.get_content_type() == "text/plain":
                email_body = part.get_payload(decode=True).decode()
                break
    else:
        email_body = email_message.get_payload(decode=True).decode()

    server.logout()
    email_subject = email_message['Subject']
    from_address = email_message['From']
    to_address = email_message['To']
    print(from_address)
    print(to_address)
    # Use ChatGPT API to generate the reply
    response = openai_client.chat.completions.create(
        model="gpt-4o",
        messages=[
            {"role": "user", "content": "You are a professional email writer"},
            {"role": "assistant", "content": "Ok"},
            {"role": "user", "content": f"Create a reply to this email:\n{email_body}"}
        ]
    )

    reply_body = response.choices[0].message.content

    # Create the email reply
    msg = MIMEMultipart()
    msg['From'] = email_account
    msg['To'] = from_address
    msg['Subject'] = f"Re: {email_subject}"
    msg.attach(MIMEText(reply_body, 'plain'))

    # Send the email using SMTP
    try:
        with smtplib.SMTP('smtp.office365.com', 587) as server:
            server.starttls()
            server.login(email_account, email_pwd)
            server.send_message(msg)
            print("Reply sent successfully.")
    except Exception as e:
        print(f"Failed to send reply: {e}")

# Generate Reply button
button = tk.Button(root, text="Generate Reply", command=reply)
button.pack()

root.mainloop()

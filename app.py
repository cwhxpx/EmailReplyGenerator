from openai import OpenAI
import win32com.client
import tkinter as tk
import config
import imapclient
import pyzmail
import email
from email.policy import default

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
    # 处理并显示邮件
    for msgid, data in response.items():
        email_message = email.message_from_bytes(data[b'BODY[]'], policy=default)
        subject = email_message['subject']
        from_ = email_message['from']
        to = email_message['to']
        emails.append(f"Subject: {subject}, From: {from_}, To: {to}")
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
root.mainloop()

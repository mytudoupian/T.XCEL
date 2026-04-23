import imaplib
import email
import smtplib
from email.mime.text import MIMEText
from email.header import Header
import time
import re
import hashlib

import os

# 从环境变量读取敏感信息
IMAP_SERVER = "imap.qq.com"
SMTP_SERVER = "smtp.qq.com"
EMAIL_ADDRESS = os.environ.get("EMAIL_ADDRESS")
EMAIL_PASSWORD = os.environ.get("EMAIL_PASSWORD")


# 激活码生成函数（请换成你自己的算法）
def generate_activation_code(machine_code):
    """根据机器码生成激活码的算法（示例为 HMAC-SHA256）"""
    secret = "my-very-secret-key-2024"
    signature = hmac.new(secret.encode(), machine_code.encode(), hashlib.sha256).hexdigest()
    code = signature[:16].upper()
    return '-'.join([code[i:i+4] for i in range(0, 16, 4)])

# ========== 核心逻辑（无需修改） ==========
def extract_machine_code(text):
    """从邮件正文提取机器码（可根据实际情况调整正则）"""
    # 示例：匹配类似 "MC-12345678" 或纯数字字母组合
    match = re.search(r"机器码[:：]?\s*([A-Z0-9\-]+)", text, re.IGNORECASE)
    if match:
        return match.group(1).strip()
    return None

def check_and_reply():
    # 1. 连接 IMAP 收取邮件
    mail = imaplib.IMAP4_SSL(IMAP_SERVER)
    mail.login(EMAIL_ADDRESS, EMAIL_PASSWORD)

    # 选中收件箱，并打印详细状态
    status, data = mail.select("INBOX")
    print(f"📬 Select INBOX 结果: 状态={status}, 数据={data}")
    if status != "OK":
        print("❌ 无法选中收件箱，请检查 163 邮箱权限或网络")
        mail.logout()
        return

    # 2. 搜索未读邮件
    status, data = mail.search(None, "UNSEEN")
    if status != "OK":
        print("未找到未读邮件")
        return

    for num in data[0].split():
        status, msg_data = mail.fetch(num, "(RFC822)")
        if status != "OK":
            continue

        msg = email.message_from_bytes(msg_data[0][1])
        from_addr = email.utils.parseaddr(msg.get("From"))[1]
        subject = msg.get("Subject", "")
        if not subject:
            for part in msg.walk():
                if part.get_content_type() == "text/plain":
                    body = part.get_payload(decode=True).decode()
                    subject = body
                    break

        print(f"收到来自 {from_addr} 的邮件，内容：{subject[:50]}...")

        # 3. 提取机器码并生成激活码
        machine_code = extract_machine_code(subject)
        if not machine_code:
            reply_text = "未能识别机器码，请检查格式后重发。"
        else:
            activation = generate_activation_code(machine_code)
            reply_text = f"您的机器码：{machine_code}\n激活码：{activation}\n感谢使用！"

        # 4. 发送回复邮件
        msg_reply = MIMEText(reply_text, "plain", "utf-8")
        msg_reply["From"] = EMAIL_ADDRESS
        msg_reply["To"] = from_addr
        msg_reply["Subject"] = Header("自动回复：您的激活码", "utf-8")

        with smtplib.SMTP_SSL(SMTP_SERVER) as smtp:
            smtp.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
            smtp.send_message(msg_reply)

        print(f"已回复 {from_addr}")

        # 5. 将邮件标记为已读
        mail.store(num, "+FLAGS", "\\Seen")

    mail.close()
    mail.logout()

if __name__ == "__main__":
    check_and_reply()

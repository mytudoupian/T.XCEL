import imaplib
import email
import smtplib
from email.mime.text import MIMEText
from email.header import Header
import time
import re
import hmac
import hashlib

import os

# 从环境变量读取敏感信息
IMAP_SERVER = "imap.qq.com"
SMTP_SERVER = "smtp.qq.com"
EMAIL_ADDRESS = os.environ.get("EMAIL_ADDRESS")
EMAIL_PASSWORD = os.environ.get("EMAIL_PASSWORD")
RSA_PUBLICKEY=os.environ.get("RSA_PUBLICKEY")
RSA_PRIVATEKEY2=os.environ.get("RSA_PRIVATEKEY2")

# 激活码生成函数（请换成你自己的算法）
def generate_activation_code(machine_code):
    """根据机器码生成激活码的算法（示例为 HMAC-SHA256）"""
    secret = "my-very-secret-key-2024"
    signature1 = hmac.new(secret.encode(), machine_code.encode(), hashlib.sha256).hexdigest()
    signature2 = hmac.new(RSA_Publickey, RSA_PRIVATEKEY2, hashlib.sha256).hexdigest()
    code = signature1[:32].upper()+signature2[:32].upper()
    return '-'.join([code[i:i+4] for i in range(0, 64, 4)])

# ========== 核心逻辑（无需修改） ==========
def extract_machine_code(text):
    # 移除不可见字符（只保留 ASCII 可见字符，避免零宽字符、控制符等干扰匹配）
    cleaned_text = re.sub(r'[^\x20-\x7e]', '', text)
    
    # 匹配 T.XCEL Machine Code=xxxx// ，xxxx 为 64-256 位十六进制字符
    pattern = r"T\.XCEL\s+Machine\s+Code=([0-9A-Fa-f]{64,256})//"
    match = re.search(pattern, cleaned_text)
    if match:
        return match.group(1).strip()
    return None

def check_and_reply():
    # 连接邮箱
    mail = imaplib.IMAP4_SSL(IMAP_SERVER)
    mail.login(EMAIL_ADDRESS, EMAIL_PASSWORD)

    status, data = mail.select("INBOX")
    print(f"📬 Select INBOX 结果: 状态={status}, 数据={data}")
    if status != "OK":
        print("❌ 无法选中收件箱")
        mail.logout()
        return

    # 搜索未读邮件
    status, data = mail.search(None, "UNSEEN")
    if status != "OK":
        print("未找到未读邮件")
        return

    max_replies = 5
    reply_count = 0

    for num in data[0].split():
        if reply_count >= max_replies:
            print(f"已达到单次最大回复数 {max_replies}，剩余邮件将在下次处理")
            break

        status, msg_data = mail.fetch(num, "(RFC822)")
        if status != "OK":
            continue

        msg = email.message_from_bytes(msg_data[0][1])
        from_addr = email.utils.parseaddr(msg.get("From"))[1]
        subject = msg.get("Subject", "")

        # ✅ 主题匹配：忽略大小写和所有空格，以 "T.XCEL Request For a key" 开头
        required_prefix = "t.xcelrequestforakey"
        cleaned_subject = subject.replace(" ", "").lower()
        if not cleaned_subject.startswith(required_prefix):
            print(f"⏭️  主题不匹配 ({from_addr}): {subject[:50]}，跳过并标记已读")
            mail.store(num, "+FLAGS", "\\Seen")
            continue

        # 获取邮件正文（纯文本）
        body_text = ""
        if msg.is_multipart():
            for part in msg.walk():
                if part.get_content_type() == "text/plain":
                    payload = part.get_payload(decode=True)
                    if payload:
                        body_text = payload.decode(errors='ignore')
                        break
        else:
            payload = msg.get_payload(decode=True)
            if payload:
                body_text = payload.decode(errors='ignore')

        # 正文前300字符检查
        preview = body_text[:300]
        print(f"收到来自 {from_addr} 的邮件，主题正确，正文前300字符已提取")

        # 提取机器码（正则已忽略大小写）
        machine_code = extract_machine_code(preview)
        if not machine_code:
            print(f"  ⏭️  未找到有效机器码，跳过并标记已读")
            mail.store(num, "+FLAGS", "\\Seen")
            continue

        # 生成激活码并回复
        activation = generate_activation_code(machine_code)
        reply_text = f"您的机器码：{machine_code}\n激活码：{activation}\n感谢使用！"

        try:
            msg_reply = MIMEText(reply_text, "plain", "utf-8")
            msg_reply["From"] = EMAIL_ADDRESS
            msg_reply["To"] = from_addr
            msg_reply["Subject"] = Header("[T.XCEL自动回复]：您的激活码", "utf-8")

            with smtplib.SMTP_SSL(SMTP_SERVER) as smtp:
                smtp.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
                smtp.send_message(msg_reply)

            print(f"  ✅ 已回复 {from_addr}")
            reply_count += 1
        except Exception as e:
            print(f"  ❌ 回复失败 ({from_addr}): {e}")
        finally:
            mail.store(num, "+FLAGS", "\\Seen")

    mail.close()
    mail.logout()
    print(f"本次共处理邮件，成功回复 {reply_count} 封")
    print(f"本次共处理邮件，成功回复 {reply_count} 封")

if __name__ == "__main__":
    check_and_reply()

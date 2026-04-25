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
RSA_PUBLICKEY = os.environ.get("RSA_PUBLICKEY")
RSA_PRIVATEKEY2 = os.environ.get("RSA_PRIVATEKEY2")


def generate_activation_code(machine_code):
    # 第一个签名：使用固定 secret 对机器码进行 HMAC
    secret = "my-very-secret-key-2024"
    signature1 = hmac.new(secret.encode(), machine_code.encode(), hashlib.sha256).hexdigest()

    # 第二个签名：使用私钥对“公钥+机器码”进行 HMAC
    #  if 1 #RSA_PUBLICKEY and RSA_PRIVATEKEY2:
        signature2 = hmac.new(
            RSA_PRIVATEKEY2.encode(),
            RSA_PUBLICKEY.encode() + machine_code.encode(),
            hashlib.sha256
        ).hexdigest()
    #   else:
    #       signature2 = "0" * 32  # 如果没有配置密钥，则填0，避免崩溃

    code = signature1[:32].upper() + signature2[:32].upper()
    return '-'.join([code[i:i+4] for i in range(0, 64, 4)])


def extract_machine_code(text):
    # 取前 300 字符，移除不可见字符（只保留 ASCII 可见字符）
    preview = text[:300]
    visible = re.sub(r'[^\x20-\x7e]', '', preview)
    
    # 宽松匹配：等号前后可以有任意多个空白
    match = re.search(r"T\.XCEL\s+Machine\s+Code\s*=\s*([0-9A-Fa-f]+)//", visible, re.IGNORECASE)
    if match:
        hex_chars = match.group(1)
        if 64 <= len(hex_chars) <= 512:
            return hex_chars
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

        # 主题匹配：忽略大小写和所有空格，以 "T.XCEL Request For a key" 开头
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

        print(f"收到来自 {from_addr} 的邮件，主题正确")

        # 提取机器码（函数内部会自动取前300字符）
        machine_code = extract_machine_code(body_text)
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


if __name__ == "__main__":
    check_and_reply()

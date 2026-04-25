import imaplib
import email
import smtplib
from email.mime.text import MIMEText
from email.header import Header
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
    secret = "my-very-secret-key-2024"
    signature1 = hmac.new(secret.encode(), machine_code.encode(), hashlib.sha256).hexdigest()

    if RSA_PUBLICKEY and RSA_PRIVATEKEY2:
        signature2 = hmac.new(
            RSA_PRIVATEKEY2.encode(),
            RSA_PUBLICKEY.encode() + machine_code.encode(),
            hashlib.sha256
        ).hexdigest()
    else:
        signature2 = "0" * 32

    code = signature1[:32].upper() + signature2[:32].upper()
    return '-'.join([code[i:i+4] for i in range(0, 64, 4)])


def extract_machine_code(text):
    # 只取前300字符
    preview = text[:300]
    # 去除不可见字符，只保留 ASCII 可见字符
    visible = re.sub(r'[^\x20-\x7e]', '', preview)

    print("========== 调试信息 ==========")
    print(f"原始文本前300字符:\n{repr(preview[:300])}")
    print(f"清理后文本前300字符:\n{repr(visible[:300])}")
    print("==============================")

    # 允许 // 前有任意空白
    match = re.search(r"T\.XCEL\s+Machine\s+Code\s*=\s*([0-9A-Fa-f]+)\s*//", visible, re.IGNORECASE)
    if match:
        hex_chars = match.group(1)
        print(f"匹配到的十六进制串长度: {len(hex_chars)}")
        if 64 <= len(hex_chars) <= 512:
            return hex_chars
        else:
            print(f"长度不在64~512之间，实际长度: {len(hex_chars)}")
    else:
        print("正则未匹配到任何内容！")
    return None


def check_and_reply():
    mail = imaplib.IMAP4_SSL(IMAP_SERVER)
    mail.login(EMAIL_ADDRESS, EMAIL_PASSWORD)

    status, data = mail.select("INBOX")
    print(f"📬 Select INBOX 结果: 状态={status}, 数据={data}")
    if status != "OK":
        print("❌ 无法选中收件箱")
        mail.logout()
        return

    status, data = mail.search(None, "UNSEEN")
    if status != "OK":
        print("未找到未读邮件")
        return

    max_replies = 5
    reply_count = 0

    for num in data[0].split():
        if reply_count >= max_replies:
            break

        status, msg_data = mail.fetch(num, "(RFC822)")
        if status != "OK":
            continue

        msg = email.message_from_bytes(msg_data[0][1])
        from_addr = email.utils.parseaddr(msg.get("From"))[1]
        subject = msg.get("Subject", "")

        required_prefix = "t.xcelrequestforakey"
        if not subject.replace(" ", "").lower().startswith(required_prefix):
            print(f"⏭️  主题不匹配 ({from_addr}): {subject[:50]}")
            mail.store(num, "+FLAGS", "\\Seen")
            continue

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

        print(f"👁️ 处理来自 {from_addr} 的邮件，主题正确")
        machine_code = extract_machine_code(body_text)
        if not machine_code:
            print(f"❌ 未找到有效机器码，跳过")
            mail.store(num, "+FLAGS", "\\Seen")
            continue

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
            print(f"✅ 已回复 {from_addr}")
            reply_count += 1
        except Exception as e:
            print(f"❌ 回复失败 ({from_addr}): {e}")
        finally:
            mail.store(num, "+FLAGS", "\\Seen")

    mail.close()
    mail.logout()
    print(f"本次共处理邮件，成功回复 {reply_count} 封")


if __name__ == "__main__":
    check_and_reply()

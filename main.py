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


if not RSA_PUBLICKEY or not RSA_PRIVATEKEY2:
    print("⚠️ 警告：RSA 密钥未设置！请检查 GitHub Secrets 中是否添加了 RSA_PUBLICKEY 和 RSA_PRIVATEKEY2。")

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
    # 先取前 2000 字符，确保完整包含机器码
    text = text[:2000]
    # 剔除所有 HTML 标签
    text_without_tags = re.sub(r'<[^>]+>', '', text)
    # 移除不可见字符，只保留 ASCII 可见字符
    visible = re.sub(r'[^\x20-\x7e]', '', text_without_tags)

    # ---------- 调试信息 ----------
    print("========== 调试信息 ==========")
    print(f"去除HTML标签后文本前500字符:\n{visible[:500]}")
    print("==============================")
    # -----------------------------

    # 宽松匹配：允许等号前后空格，连续十六进制字符直至 //
    match = re.search(r"T\.XCEL\s+Machine\s+Code\s*=\s*([0-9A-Fa-f]+)//", visible, re.IGNORECASE)
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

        # ---------- 改进的正文提取 ----------
        body_text = ""
        html_body = ""
        if msg.is_multipart():
            for part in msg.walk():
                content_type = part.get_content_type()
                if content_type == "text/plain":
                    payload = part.get_payload(decode=True)
                    if payload:
                        body_text = payload.decode(errors='ignore')
                        break
                elif content_type == "text/html":
                    payload = part.get_payload(decode=True)
                    if payload and not html_body:
                        html_body = payload.decode(errors='ignore')
        else:
            payload = msg.get_payload(decode=True)
            if payload:
                content_type = msg.get_content_type()
                if content_type == "text/plain":
                    body_text = payload.decode(errors='ignore')
                elif content_type == "text/html":
                    html_body = payload.decode(errors='ignore')
        
        # 如果没有纯文本，则从 HTML 中提取纯文本
        if not body_text and html_body:
            body_text = re.sub(r'<[^>]+>', '', html_body)
        # -----------------------------------

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

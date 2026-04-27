import imaplib
import email
import smtplib
from email.mime.text import MIMEText
from email.header import Header
import re
import hashlib
import os

# ==================== 常量（与 VBA 端完全一致） ====================
SHARED_KEY_HEX = "A1B2"                      # 2 字节共享密钥
SHARED_KEY_BYTES = bytes.fromhex(SHARED_KEY_HEX)
SERVER_SECRET = "T.XCEL-Server-Secret-2026"  # 激活码生成密钥

# ==================== LCG 伪随机流生成器（与 VBA 完全一致） ====================
def lcg_generate_stream(seed_bytes: bytes, length: int) -> bytes:
    """seed_bytes 必须为 6 字节（key 2 + rand 4）"""
    if len(seed_bytes) < 6:
        raise ValueError("种子必须为 6 字节")
    m = 2**32
    a = 1103515245
    c = 12345
    # 大端整数构造并取模得到初始状态
    big_num = 0
    for b in seed_bytes[:6]:
        big_num = big_num * 256 + b
    state = big_num % m
    stream = bytearray()
    for _ in range(length):
        state = (state * a + c) % m
        stream.append(state & 0xFF)
    return bytes(stream)

# ==================== 机器码解密与验证 ====================
def verify_and_extract_base(machine_code: str):
    """验证 72 位机器码，成功则返回 (base_hex, rand_hex)，否则返回 (None, None)"""
    if len(machine_code) != 72:
        return None, None

    cipher_hex = machine_code[:64]
    expected_checksum = machine_code[64:].lower()

    # 1. 校验和验证
    actual_checksum = hashlib.md5(cipher_hex.encode()).hexdigest()[:8]
    if actual_checksum != expected_checksum:
        print("校验和不匹配")
        return None, None

    try:
        cipher_bytes = bytes.fromhex(cipher_hex)  # 32 bytes
    except ValueError:
        return None, None

    rand_bytes = cipher_bytes[:4]          # 随机数（4 字节）
    enc_base  = cipher_bytes[4:20]         # 加密基础码（16 字节）
    tail      = cipher_bytes[20:32]        # 流尾部（12 字节）

    # 2. 重建种子并生成流
    seed = SHARED_KEY_BYTES + rand_bytes   # 6 字节
    stream = lcg_generate_stream(seed, 28) # 28 字节

    # 3. 尾部验证
    if stream[16:] != tail:
        print("尾部验证失败")
        return None, None

    # 4. 解密基础码
    base_bytes = bytes(a ^ b for a, b in zip(enc_base, stream[:16]))
    base_hex = base_bytes.hex()
    rand_hex = rand_bytes.hex()
    return base_hex, rand_hex

# ==================== 激活码生成（40 位） ====================
def generate_activation_code(base_hex: str, rand_hex: str) -> str:
    """base_hex: 32 位固定基础码, rand_hex: 8 位随机数"""
    hash_input = base_hex + SERVER_SECRET
    active_hash = hashlib.md5(hash_input.encode()).hexdigest()[:32]
    return active_hash + rand_hex

# ==================== 邮件正文中提取机器码 ====================
def extract_machine_code(text: str):
    """从文本中提取符合格式的 72 位机器码"""
    # 先取前 2000 字符，剔除 HTML 标签，保留可见字符
    text = text[:2000]
    text_without_tags = re.sub(r'<[^>]+>', '', text)
    visible = re.sub(r'[^\x20-\x7e]', '', text_without_tags)

    match = re.search(r"T\.XCEL\s+Machine\s+Code\s*=\s*([0-9A-Fa-f]{72})//", visible, re.IGNORECASE)
    if match:
        return match.group(1).strip()
    return None

# ==================== 邮箱配置（从环境变量读取） ====================
IMAP_SERVER = "imap.qq.com"
SMTP_SERVER = "smtp.qq.com"
EMAIL_ADDRESS = os.environ.get("EMAIL_ADDRESS")
EMAIL_PASSWORD = os.environ.get("EMAIL_PASSWORD")

# ==================== 邮件检查与自动回复 ====================
def check_and_reply():
    if not EMAIL_ADDRESS or not EMAIL_PASSWORD:
        print("❌ 环境变量 EMAIL_ADDRESS / EMAIL_PASSWORD 未设置")
        return

    mail = imaplib.IMAP4_SSL(IMAP_SERVER)
    mail.login(EMAIL_ADDRESS, EMAIL_PASSWORD)

    status, data = mail.select("INBOX")
    print(f"📬 Select INBOX 结果: 状态={status}, 数据={data}")
    if status != "OK":
        print("❌ 无法选中收件箱")
        mail.logout()
        return

    status, data = mail.search(None, "UNSEEN")
    if status != "OK" or not data[0]:
        mail.close()
        mail.logout()
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

        # 主题过滤
        required_prefix = "t.xcelrequestforakey"
        if not subject.replace(" ", "").lower().startswith(required_prefix):
            print(f"⏭️  主题不匹配 ({from_addr}): {subject[:50]}")
            mail.store(num, "+FLAGS", "\\Seen")
            continue

        # 提取正文（优先 text/plain，否则从 text/html 剥离）
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

        if not body_text and html_body:
            body_text = re.sub(r'<[^>]+>', '', html_body)

        print(f"👁️ 处理来自 {from_addr} 的邮件，主题正确")

        # 提取 72 位机器码
        raw_code = extract_machine_code(body_text)
        if not raw_code:
            print("❌ 未找到有效机器码，跳过")
            mail.store(num, "+FLAGS", "\\Seen")
            continue

        # 验证并解密机器码，获取 base 和 rand
        base_hex, rand_hex = verify_and_extract_base(raw_code)
        if not base_hex:
            print("❌ 机器码验证失败，可能被篡改")
            mail.store(num, "+FLAGS", "\\Seen")
            continue

        # 生成 40 位激活码
        activation = generate_activation_code(base_hex, rand_hex)
        reply_text = f"👀您的机器码：{raw_code}\n激活码：{activation}\n🎲感谢使用！"

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

#!/usr/bin/env python3
"""
Microsoft邮件处理脚本
用于收发Microsoft账号的邮件
"""
import os

from xml.etree.ElementTree import fromstring, tostringlist
import requests
import logging
from datetime import datetime
from typing import Dict, List
import configparser
import winreg
import time

import asyncio
from aiosmtpd.controller import Controller
from aiosmtpd.smtp import SMTP

from email import message_from_bytes
from email.header import decode_header

# def get_proxy():
#     try:
#         with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Internet Settings") as key:
#             proxy_enable, _ = winreg.QueryValueEx(key, "ProxyEnable")
#             proxy_server, _ = winreg.QueryValueEx(key, "ProxyServer")
            
#             if proxy_enable and proxy_server:
#                 proxy_parts = proxy_server.split(":")
#                 if len(proxy_parts) == 2:
#                     return {"http": f"http://{proxy_server}", "https": f"http://{proxy_server}"}
#     except WindowsError:
#         pass
#     return {"http": None, "https": None}

def get_proxy():
    # Environment variables can be lowercase or uppercase
    http_proxy = os.environ.get("http_proxy") or os.environ.get("HTTP_PROXY")
    https_proxy = os.environ.get("https_proxy") or os.environ.get("HTTPS_PROXY")
    
    # Clean up formatting to ensure they start with http:// or https://
    def clean_scheme(url, scheme):
        if url and not url.startswith(("http://", "https://", "socks://")):
            return f"{scheme}://{url}"
        return url

    return {
        "http": clean_scheme(http_proxy, "http"),
        "https": clean_scheme(https_proxy, "http") # Or "https" depending on proxy setup
    }

def load_config():
    """从config.txt加载配置"""
    config = configparser.ConfigParser()
    config.read('config.txt', encoding='utf-8')
    return config

def save_config(config):
    """保存配置到config.txt"""
    with open('config.txt', 'w', encoding='utf-8') as f:
        config.write(f)

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

config = load_config()
microsoft_config = config['microsoft']

CLIENT_ID = microsoft_config['client_id']
GRAPH_API_ENDPOINT = 'https://graph.microsoft.com/v1.0'
TOKEN_URL = 'https://login.microsoftonline.com/common/oauth2/v2.0/token'

class EmailClient:
    def __init__(self):
        config = load_config()
        if not config.has_section('tokens'):
            config.add_section('tokens')
        self.config = config
        self.refresh_token = config['tokens'].get('refresh_token', '')
        self.access_token = config['tokens'].get('access_token', '')
        expires_at_str = config['tokens'].get('expires_at', '1970-01-01 00:00:00')
        self.expires_at = datetime.strptime(expires_at_str, '%Y-%m-%d %H:%M:%S').timestamp()
    
    def is_token_expired(self) -> bool:
        """检查access token是否过期或即将过期"""
        buffer_time = 300
        return datetime.now().timestamp() + buffer_time >= self.expires_at
    
    def refresh_access_token(self) -> None:
        """刷新访问令牌"""
        refresh_params = {
            'client_id': CLIENT_ID,
            'refresh_token': self.refresh_token,
            'grant_type': 'refresh_token',
        }
        
        try:
            response = requests.post(TOKEN_URL, data=refresh_params, proxies=get_proxy())
            response.raise_for_status()
            tokens = response.json()
            
            self.access_token = tokens['access_token']
            self.expires_at = time.time() + tokens['expires_in']
            expires_at_str = datetime.fromtimestamp(self.expires_at).strftime('%Y-%m-%d %H:%M:%S')
            
            self.config['tokens']['access_token'] = self.access_token
            self.config['tokens']['expires_at'] = expires_at_str
            
            if 'refresh_token' in tokens:
                self.refresh_token = tokens['refresh_token']
                self.config['tokens']['refresh_token'] = self.refresh_token
            save_config(self.config)
        except requests.RequestException as e:
            logger.error(f"刷新访问令牌失败: {e}")
            raise

    def ensure_token_valid(self):
        """确保token有效"""
        if not self.access_token or self.is_token_expired():
            self.refresh_access_token()

    def get_messages(self, folder_id: str = 'inbox', top: int = 10) -> List[Dict]:
        """获取指定文件夹的邮件
        
        Args:
            folder_id: 文件夹ID, 默认为'inbox'
            top: 获取的邮件数量
        """
        self.ensure_token_valid()
        
        headers = {
            'Authorization': f'Bearer {self.access_token}',
            'Accept': 'application/json',
            'Prefer': 'outlook.body-content-type="text"'
        }
        
        query_params = {
            '$top': top,
            '$select': 'subject,receivedDateTime,from,body',
            '$orderby': 'receivedDateTime DESC'
        }
        
        try:
            response = requests.get(
                f'{GRAPH_API_ENDPOINT}/me/mailFolders/{folder_id}/messages',
                headers=headers,
                params=query_params,
                proxies=get_proxy()
            )
            response.raise_for_status()
            return response.json()['value']
        except requests.RequestException as e:
            logger.error(f"获取邮件失败: {e}")
            if response.status_code == 401:
                self.refresh_access_token()
                return self.get_messages(folder_id, top)
            raise

    def get_junk_messages(self, top: int = 10) -> List[Dict]:
        """获取垃圾邮件文件夹中的邮件"""
        return self.get_messages(folder_id='junkemail', top=top)

    def send_email(self, to_recipients: List[str], subject: str, content: str, is_html: bool = False) -> bool:
        """发送邮件
        
        Args:
            to_recipients: 收件人邮箱地址列表
            subject: 邮件主题
            content: 邮件内容
            is_html: 内容是否为HTML格式，默认为False
            
        Returns:
            bool: 发送是否成功
        """
        self.ensure_token_valid()
        
        headers = {
            'Authorization': f'Bearer {self.access_token}',
            'Content-Type': 'application/json'
        }
        
        email_msg = {
            'message': {
                'subject': subject,
                'body': {
                    'contentType': 'HTML' if is_html else 'Text',
                    'content': content
                },
                'toRecipients': [
                    {
                        'emailAddress': {
                            'address': recipient
                        }
                    } for recipient in to_recipients
                ]
            }
        }
        
        try:
            response = requests.post(
                f'{GRAPH_API_ENDPOINT}/me/sendMail',
                headers=headers,
                json=email_msg,
                proxies=get_proxy()
            )
            response.raise_for_status()
            logger.info(f"邮件已成功发送给 {', '.join(to_recipients)}")
            return True
        except requests.RequestException as e:
            logger.error(f"发送邮件失败: {e}")
            if response.status_code == 401:
                self.refresh_access_token()
                return self.send_email(to_recipients, subject, content, is_html)
            raise

# Optional: Enable logging to see connection details in the console
logging.basicConfig(level=logging.INFO)

class CustomSMTPHandler:
    async def handle_DATA(self, server, session, envelope):
        print('--- New Message Received ---')
        # 1. Parse the raw bytes into an EmailMessage object
        msg = message_from_bytes(envelope.content)
        
        # 2. Extract the Subject header
        raw_subject = msg.get('Subject', '(No Subject)')
        
        # 3. Decode the subject safely (handles UTF-8, Base64, etc.)
        subject_parts = decode_header(raw_subject)
        decoded_subject = ""
        for part, encoding in subject_parts:
            if isinstance(part, bytes):
                decoded_subject += part.decode(encoding or 'utf-8', errors='replace')
            else:
                decoded_subject += part
         
        # 4. Get the primary (overall) Content-Type of the email
        # .get_content_type() normalizes it to lowercase (e.g., 'multipart/alternative')
        overall_content_type = msg.get_content_type()
        print(f"Overall Email Content-Type: {overall_content_type}")
        
        # 3. Inspect individual parts if the email is multipart
        if msg.is_multipart():
            print("\n--- Document Structure ---")
            for i, part in enumerate(msg.walk(), start=1):
                part_type = part.get_content_type()
                part_filename = part.get_filename()
                
                print(f"Part {i}: {part_type}")
                if part_filename:
                    print(f"   -> Attachment Filename: {part_filename}")
        else:
            print("This is a single-part simple email.")
            
        body_text = ""
        body_html = ""
        
        # 4. Walk through the parts of the email
        if msg.is_multipart():
            
            # todo: uncheck 
            for part in msg.walk():
                content_type = part.get_content_type()
                content_disposition = str(part.get("Content-Disposition"))

                # Skip attachments
                if "attachment" in content_disposition:
                    continue

                # Extract text or HTML body
                if content_type == "text/plain":
                    # decode=True handles Base64/Quoted-Printable transfer encodings automatically
                    body_text = part.get_payload(decode=True).decode(part.get_content_charset() or 'utf-8', errors='replace')
                elif content_type == "text/html":
                    body_html = part.get_payload(decode=True).decode(part.get_content_charset() or 'utf-8', errors='replace')
                    #body_text = part.get_payload(decode=True).decode(part.get_content_charset() or 'utf-8', errors='replace')
        else:
            # If it's not multipart, the message itself is the body
            body_text = msg.get_payload(decode=True).decode(msg.get_content_charset() or 'utf-8', errors='replace')
            
        print('--- New Message Received ---')
        print(f"Subject: {decoded_subject}")
        print(f'From: {envelope.mail_from}')
        print(f'To:   {envelope.rcpt_tos}')  
        print(f"Content-Type:{overall_content_type}")
        
        print('-' * 30)
        print(f"Content:{body_text}")
        print('-' * 30)       
        print(f"Content:{body_html}")
        print('-' * 30)    
               
        client = EmailClient()
        recipients = envelope.rcpt_tos  # 替换为实际的收件人邮箱
        print("\n发送邮件:")
        subject = decoded_subject       # 替换为实际发送邮件的主题
        if msg.is_multipart():  
             # todo: uncheck  
            if client.send_email(recipients, subject, body_text,  ):
                print("邮件发送成功！")
        else:
            if client.send_email(recipients, subject, body_text, overall_content_type == "text/html" ):
                print("邮件发送成功！")

        # Return a standard SMTP 250 OK success response
        return '250 Message accepted for delivery'

async def daemon():
    # 创建控制器，它将处理服务器的启动和停止
    # handler: 指定处理邮件逻辑的类实例
    # hostname: 监听的地址 (例如 '127.0.0.1' 或 'localhost')
    # port: 监听的端口
    handler = CustomSMTPHandler()
    
    # Run the server locally on port 8025
    controller = Controller(handler, hostname='192.168.5.148', port=8025)
    controller.start()
    
    print("SMTP Server running on 192.168.5.148:8025. Press Ctrl+C to stop.")
    
    try:
        # Keep the main loop running while the controller handles connections
        while True:
            await asyncio.sleep(3600)
    except KeyboardInterrupt:
        print("\nStopping SMTP Server...")
    finally:
        controller.stop()


async def main():
    try:  
        client = EmailClient()
        # 获取收件箱邮件,top=n表示获取最新n封邮件
        messages = client.get_messages(top=3)
        print("\n收件箱最新邮件:")
        for msg in messages:
            print("\n" + "="*50)
            print(f"主题: {msg['subject']}")
            print(f"发件人: {msg['from']['emailAddress']['address']}")
            print(f"时间: {msg['receivedDateTime']}")
            print(f"\n邮件内容:{msg['body']['content']}")
            
        # 获取垃圾邮件,top=n表示获取最新n封邮件
        junk_messages = client.get_junk_messages(top=3)
        print("\n垃圾邮件文件夹最新邮件:")
        for msg in junk_messages:
            print("\n" + "="*50)
            print(f"主题: {msg['subject']}")
            print(f"发件人: {msg['from']['emailAddress']['address']}")
            print(f"时间: {msg['receivedDateTime']}")
            print(f"\n邮件内容:{msg['body']['content']}")
        await asyncio.sleep(3600)   
    except Exception as e:
        logger.error(f"程序执行出错: {e}")
        raise

if __name__ == '__main__':
    # asyncio.run(main())
    # 运行 main 协程
    asyncio.run(daemon())



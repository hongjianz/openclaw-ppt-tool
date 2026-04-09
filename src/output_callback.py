"""
输出回调模块 - PPT生成后的后处理

功能:
- 上传到飞书云空间
- 发送到飞书聊天
- 上传到其他云存储
- 自定义回调钩子
"""

import os
import json
import requests
from typing import Optional, Callable, Dict, Any


class OutputCallback:
    """输出回调处理器"""

    def __init__(self):
        self.callbacks = []

    def register(self, callback: Callable[[str], None]):
        """注册回调函数"""
        self.callbacks.append(callback)

    def execute(self, file_path: str):
        """执行所有回调"""
        for callback in self.callbacks:
            try:
                callback(file_path)
            except Exception as e:
                print(f"回调执行失败: {e}")


class FeishuUploader:
    """飞书上传器"""

    def __init__(self, app_id: str = None, app_secret: str = None):
        self.app_id = app_id or os.getenv('FEISHU_APP_ID', '')
        self.app_secret = app_secret or os.getenv('FEISHU_APP_SECRET', '')
        self.access_token = None
        self.base_url = "https://open.feishu.cn/open-apis"

    def get_access_token(self) -> str:
        """获取访问令牌"""
        if self.access_token:
            return self.access_token

        if not self.app_id or not self.app_secret:
            raise Exception("未配置飞书应用凭证")

        url = f"{self.base_url}/auth/v3/tenant_access_token/internal"
        response = requests.post(url, json={
            "app_id": self.app_id,
            "app_secret": self.app_secret
        })
        result = response.json()

        if result.get('code') != 0:
            raise Exception(f"获取Token失败: {result.get('msg')}")

        self.access_token = result['tenant_access_token']
        return self.access_token

    def upload_to_drive(self, file_path: str, folder_token: str = None) -> Dict[str, Any]:
        """
        上传文件到飞书云空间

        Args:
            file_path: 本地文件路径
            folder_token: 文件夹token(可选,默认上传到根目录)

        Returns:
            上传结果,包含file_token和url
        """
        token = self.get_access_token()

        # 1. 准备上传
        url = f"{self.base_url}/drive/v1/files/upload_prepare"
        headers = {"Authorization": f"Bearer {token}"}

        file_name = os.path.basename(file_path)
        file_size = os.path.getsize(file_path)

        prepare_data = {
            "file_name": file_name,
            "parent_type": "explorer" if not folder_token else "folder",
            "parent_node": folder_token or "",
            "size": file_size,
            "upload_scene": "cloud_space"
        }

        response = requests.post(url, headers=headers, json=prepare_data)
        prepare_result = response.json()

        if prepare_result.get('code') != 0:
            raise Exception(f"准备上传失败: {prepare_result.get('msg')}")

        upload_id = prepare_result['data']['upload_id']

        # 2. 分片上传(简化版,小文件一次性上传)
        url = f"{self.base_url}/drive/v1/files/upload_part"
        headers["Upload-ID"] = upload_id

        with open(file_path, 'rb') as f:
            file_content = f.read()

        files = {
            'file': (file_name, file_content, 'application/vnd.openxmlformats-officedocument.presentationml.presentation')
        }

        response = requests.post(url, headers=headers, files=files)
        upload_result = response.json()

        if upload_result.get('code') != 0:
            raise Exception(f"上传失败: {upload_result.get('msg')}")

        # 3. 完成上传
        url = f"{self.base_url}/drive/v1/files/upload_finish"
        finish_data = {"upload_id": upload_id}

        response = requests.post(url, headers=headers, json=finish_data)
        finish_result = response.json()

        if finish_result.get('code') != 0:
            raise Exception(f"完成上传失败: {finish_result.get('msg')}")

        file_token = finish_result['data']['file_token']
        file_url = f"https://www.feishu.cn/drive/file/{file_token}"

        return {
            'file_token': file_token,
            'file_url': file_url,
            'file_name': file_name
        }

    def send_to_chat(self, file_path: str, chat_id: str, message: str = None) -> Dict[str, Any]:
        """
        发送文件到飞书聊天

        Args:
            file_path: 本地文件路径
            chat_id: 聊天ID
            message: 附加消息

        Returns:
            发送结果
        """
        # 先上传到云空间
        upload_result = self.upload_to_drive(file_path)

        # 发送消息
        token = self.get_access_token()
        url = f"{self.base_url}/im/v1/messages"

        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json"
        }

        msg_content = {
            "msg_type": "post",
            "content": {
                "post": {
                    "zh_cn": {
                        "title": "PPT生成完成",
                        "content": [
                            [
                                {
                                    "tag": "text",
                                    "text": message or f"演示文稿已生成: {upload_result['file_name']}"
                                }
                            ],
                            [
                                {
                                    "tag": "a",
                                    "text": "点击查看文件",
                                    "href": upload_result['file_url']
                                }
                            ]
                        ]
                    }
                }
            }
        }

        params = {"receive_id_type": "chat_id"}
        payload = {
            "receive_id": chat_id,
            **msg_content
        }

        response = requests.post(url, params=params, headers=headers, json=payload)
        result = response.json()

        if result.get('code') != 0:
            raise Exception(f"发送消息失败: {result.get('msg')}")

        return result['data']


def create_upload_callback(uploader: FeishuUploader, folder_token: str = None):
    """创建上传回调函数"""
    def callback(file_path: str):
        print(f"\n正在上传到飞书云空间...")
        result = uploader.upload_to_drive(file_path, folder_token)
        print(f"✓ 上传成功!")
        print(f"  文件名: {result['file_name']}")
        print(f"  链接: {result['file_url']}")
    return callback


def create_chat_send_callback(uploader: FeishuUploader, chat_id: str, message: str = None):
    """创建聊天发送回调函数"""
    def callback(file_path: str):
        print(f"\n正在发送到飞书聊天...")
        result = uploader.send_to_chat(file_path, chat_id, message)
        print(f"✓ 发送成功!")
    return callback

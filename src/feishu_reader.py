"""
飞书文档读取模块 - 从飞书云文档读取内容

功能:
- 通过飞书开放API读取文档内容
- 支持云文档和知识库文档
- 自动转换为Markdown格式

注意: 需要配置飞书应用凭证
"""

import os
import json
import requests
from typing import Optional


class FeishuDocReader:
    """飞书文档读取器"""

    def __init__(self, app_id: str = None, app_secret: str = None):
        """
        初始化飞书读取器

        Args:
            app_id: 飞书应用ID
            app_secret: 飞书应用密钥
        """
        self.app_id = app_id or os.getenv('FEISHU_APP_ID', '')
        self.app_secret = app_secret or os.getenv('FEISHU_APP_SECRET', '')
        self.access_token = None
        self.base_url = "https://open.feishu.cn/open-apis"

    def get_access_token(self) -> str:
        """获取访问令牌"""
        if self.access_token:
            return self.access_token

        if not self.app_id or not self.app_secret:
            raise Exception(
                "未配置飞书应用凭证\n"
                "请设置环境变量 FEISHU_APP_ID 和 FEISHU_APP_SECRET\n"
                "或在代码中传入 app_id 和 app_secret"
            )

        url = f"{self.base_url}/auth/v3/tenant_access_token/internal"
        payload = {
            "app_id": self.app_id,
            "app_secret": self.app_secret
        }

        response = requests.post(url, json=payload)
        result = response.json()

        if result.get('code') != 0:
            raise Exception(f"获取Token失败: {result.get('msg')}")

        self.access_token = result['tenant_access_token']
        return self.access_token

    def read_document(self, doc_url_or_token: str) -> str:
        """
        读取飞书文档内容

        Args:
            doc_url_or_token: 文档URL或document_token

        Returns:
            Markdown格式的文档内容
        """
        # 提取document_token
        if 'feishu.cn' in doc_url_or_token or 'larksuite.com' in doc_url_or_token:
            doc_token = self._extract_doc_token(doc_url_or_token)
        else:
            doc_token = doc_url_or_token

        if not doc_token:
            raise Exception("无法解析文档Token")

        # 获取访问令牌
        token = self.get_access_token()

        # 读取文档内容
        url = f"{self.base_url}/docx/v1/documents/{doc_token}/raw_content"
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json"
        }

        params = {
            "lang": 0  # 0=简体中文
        }

        response = requests.get(url, headers=headers, params=params)
        result = response.json()

        if result.get('code') != 0:
            raise Exception(f"读取文档失败: {result.get('msg')}")

        # 提取内容
        content = result.get('data', {}).get('content', '')
        return content

    def _extract_doc_token(self, url: str) -> Optional[str]:
        """从URL中提取document_token"""
        # 示例URL: https://xxx.feishu.cn/docx/xxxxx
        parts = url.split('/')
        for i, part in enumerate(parts):
            if part in ['docx', 'docs', 'wiki'] and i + 1 < len(parts):
                token = parts[i + 1].split('?')[0]
                return token
        return None


# 简化用法: 直接传入文本
def create_ppt_from_text(content: str, output_path: str, **kwargs):
    """
    从文本内容直接生成PPT(无需文件)

    Args:
        content: 文稿内容(Markdown或纯文本)
        output_path: 输出PPT路径
        **kwargs: 其他参数传递给main函数

    Example:
        from src.feishu_reader import create_ppt_from_text

        content = '''
        # 我的演示文稿

        ## 副标题

        ---

        # 第一页

        - 要点1
        - 要点2
        '''

        create_ppt_from_text(content, 'output.pptx')
    """
    import sys
    import tempfile

    # 创建临时文件
    with tempfile.NamedTemporaryFile(mode='w', suffix='.md', delete=False, encoding='utf-8') as f:
        f.write(content)
        temp_file = f.name

    try:
        # 构建命令行参数
        sys.argv = ['main.py', '-i', temp_file, '-o', output_path]

        # 添加其他参数
        for key, value in kwargs.items():
            if isinstance(value, bool):
                if value:
                    sys.argv.append(f'--{key.replace("_", "-")}')
            else:
                sys.argv.append(f'--{key.replace("_", "-")}')
                sys.argv.append(str(value))

        # 导入并运行main
        from main import main
        main()

    finally:
        # 清理临时文件
        if os.path.exists(temp_file):
            os.unlink(temp_file)

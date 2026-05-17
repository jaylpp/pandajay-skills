#!/usr/bin/env python3
"""
Paper Panda AIGC改写API调用脚本
用于降低文本的AI检测率或查重率
"""

import sys
import os
import json
import requests
from pathlib import Path
from typing import Dict, Optional, Tuple

# API配置
API_ENDPOINT = "https://paperpanda.cn/api/v1/aigc/rewrite"
MAX_LENGTH = 6000
CHUNK_SIZE = 5000
CONFIG_FILE = Path.home() / ".paperpanda_key"


def get_api_key() -> Optional[str]:
    """从环境变量或配置文件获取API Key"""
    key = os.environ.get("PAPERPANDA_API_KEY")
    if key:
        return key
    if CONFIG_FILE.exists():
        return CONFIG_FILE.read_text().strip()
    return None


def save_api_key(key: str):
    """保存API Key到配置文件"""
    CONFIG_FILE.write_text(key)
    CONFIG_FILE.chmod(0o600)


API_KEY = get_api_key()

# 改写类型映射
REWRITE_TYPES = {
    "知网降AI": 1,
    "维普降AI": 2,
    "通用降AI": 3,
    "英文知网降AI": 4,
    "通用降重": 5,
    "格子达降AI": 6,
    "双降知网": 7,
    "双降维普": 8,
    "双降通用": 9,
    "双降格子达": 10,
    "朱雀降AI": 11,
    "特价降AI": 12,
    "英文维普降AI": 13,
}

# 服务类型映射
SERVICE_TYPES = {
    "ai-reduce-value": "降AI-英文版（超值）",
    "ai-reduce": "降AI-中文版（标准）",
    "duplicate-reduce": "降重服务",
    "ai-duplicate": "降AI+降重（双重优化）",
    "advance": "高级咨询（定制化）",
}


def split_text(content: str, chunk_size: int = CHUNK_SIZE) -> list:
    """按段落分割长文本"""
    paragraphs = content.split('\n')
    chunks = []
    current_chunk = []

    for para in paragraphs:
        para = para.strip()
        if not para:
            continue

        # 如果单个段落超过chunk_size，按句子分割
        if len(para) > chunk_size:
            sentences = para.split('。')
            for sent in sentences:
                sent = sent.strip()
                if not sent:
                    continue
                if len('。'.join(current_chunk + [sent])) > chunk_size:
                    if current_chunk:
                        chunks.append('。'.join(current_chunk))
                        current_chunk = [sent]
                    else:
                        # 单个句子太长，强制分割
                        for i in range(0, len(sent), chunk_size):
                            chunks.append(sent[i:i+chunk_size])
                else:
                    current_chunk.append(sent)
        else:
            if len('\n'.join(current_chunk + [para])) > chunk_size:
                if current_chunk:
                    chunks.append('\n'.join(current_chunk))
                    current_chunk = [para]
                else:
                    chunks.append(para)
            else:
                current_chunk.append(para)

    if current_chunk:
        chunks.append('\n'.join(current_chunk))

    return chunks


def rewrite_text(
    content: str,
    rewrite_type: int = 3,
    service_type: str = "ai-reduce"
) -> Tuple[bool, Dict]:
    """
    调用AIGC改写API

    Args:
        content: 要改写的文本内容
        rewrite_type: 改写类型ID (1-13)
        service_type: 服务类型

    Returns:
        (success, result)
    """
    headers = {
        "Authorization": f"Bearer {API_KEY}",
        "Content-Type": "application/json"
    }

    data = {
        "content": content,
        "service_type": service_type,
        "rewrite_type": rewrite_type
    }

    try:
        response = requests.post(
            API_ENDPOINT,
            headers=headers,
            json=data,
            timeout=60
        )

        result = response.json()
        result["status_code"] = response.status_code

        if response.status_code == 200 and result.get("code") == 200:
            return True, result
        else:
            return False, result

    except requests.exceptions.Timeout:
        return False, {
            "error": "请求超时",
            "message": "API请求超时，请稍后重试"
        }
    except requests.exceptions.RequestException as e:
        return False, {
            "error": str(e),
            "message": f"请求失败: {str(e)}"
        }


def rewrite_long_text(
    content: str,
    rewrite_type: int = 3,
    service_type: str = "ai-reduce"
) -> Tuple[bool, Dict]:
    """
    处理长文本改写（自动分段）

    Args:
        content: 要改写的文本内容
        rewrite_type: 改写类型ID
        service_type: 服务类型

    Returns:
        (success, result)
    """
    if len(content) <= MAX_LENGTH:
        return rewrite_text(content, rewrite_type, service_type)

    # 分段处理
    chunks = split_text(content, CHUNK_SIZE)
    results = []
    total_amount = 0
    failed_chunks = []

    for i, chunk in enumerate(chunks, 1):
        success, result = rewrite_text(chunk, rewrite_type, service_type)
        if success:
            results.append(result.get("data", ""))
            total_amount += result.get("amount", 0)
        else:
            failed_chunks.append(i)
            results.append(chunk)  # 保留原文

    return len(failed_chunks) == 0, {
        "data": "\n".join(results),
        "amount": total_amount,
        "chunks_total": len(chunks),
        "chunks_failed": len(failed_chunks),
        "failed_indices": failed_chunks
    }


def get_rewrite_type_name(rewrite_type: int) -> str:
    """获取改写类型名称"""
    for name, rt_id in REWRITE_TYPES.items():
        if rt_id == rewrite_type:
            return name
    return f"未知类型({rewrite_type})"


def get_service_type_name(service_type: str) -> str:
    """获取服务类型名称"""
    return SERVICE_TYPES.get(service_type, service_type)


def detect_platform(text: str) -> Tuple[int, str]:
    """
    根据文本内容智能检测改写类型

    Returns:
        (rewrite_type, platform_name)
    """
    # 简单的关键词检测
    keywords_map = {
        "知网": (1, "知网降AI"),
        "维普": (2, "维普降AI"),
        "格子达": (6, "格子达降AI"),
        "朱雀": (11, "朱雀降AI"),
        "特价": (12, "特价降AI"),
    }

    for keyword, (rt, name) in keywords_map.items():
        if keyword in text:
            return rt, name

    # 默认返回通用降AI
    return 3, "通用降AI"


if __name__ == "__main__":
    # 设置API Key
    if len(sys.argv) >= 2 and sys.argv[1] == "--set-key":
        if len(sys.argv) < 3:
            print("用法: python rewrite.py --set-key <your_api_key>")
            sys.exit(1)
        save_api_key(sys.argv[2])
        print(f"✅ API Key 已保存到 {CONFIG_FILE}")
        sys.exit(0)

    # 检查API Key
    if not API_KEY:
        print("❌ 未配置 API Key")
        print(f"\n请通过以下方式之一配置：")
        print(f"  1. 运行: python3 {sys.argv[0]} --set-key <your_api_key>")
        print(f"  2. 设置环境变量: export PAPERPANDA_API_KEY=<your_api_key>")
        print(f"\nAPI Key 获取地址: https://paperpanda.cn (注册后在API密钥管理中创建)")
        sys.exit(1)

    if len(sys.argv) < 2:
        print("用法: python rewrite.py <文本内容> [改写类型] [服务类型]")
        print("\n改写类型:")
        for name, rt_id in REWRITE_TYPES.items():
            print(f"  {rt_id}: {name}")
        sys.exit(1)

    text = sys.argv[1]
    rt = int(sys.argv[2]) if len(sys.argv) > 2 else 3
    st = sys.argv[3] if len(sys.argv) > 3 else "ai-reduce"

    success, result = rewrite_long_text(text, rt, st)

    if success:
        print(f"✅ 改写成功")
        print(f"消费金额: {result.get('amount', 0)} 元")
        if "chunks_total" in result:
            print(f"分段处理: {result['chunks_total']}段, 失败{result['chunks_failed']}段")
        print(f"\n改写结果:\n{result['data']}")
    else:
        print(f"❌ 改写失败: {result.get('message', '未知错误')}")
        if "error" in result:
            print(f"错误详情: {result['error']}")
        sys.exit(1)

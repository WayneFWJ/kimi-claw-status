#!/usr/bin/env python3
"""
kimi-claw 状态监控脚本 - 简化版
使用 HTTP POST 直接写入 Upstash Redis
"""

import sys
import json
import time
import http.client

UPSTASH_HOST = "artistic-prawn-73537.upstash.io"
UPSTASH_TOKEN = "gQAAAAAAAR9BAAIncDIzOTVlYjkxYWZhZjU0NGNmOWEwMGE1OTEzZGZmZDE3M3AyNzM1Mzc"

def redis_request(method, path, body=None):
    """发送 Redis HTTP 请求"""
    try:
        conn = http.client.HTTPSConnection(UPSTASH_HOST, timeout=5)
        headers = {
            "Authorization": f"Bearer {UPSTASH_TOKEN}",
            "Content-Type": "application/json"
        }
        
        if body and isinstance(body, dict):
            body = json.dumps(body)
            
        conn.request(method, path, body, headers)
        response = conn.getresponse()
        data = response.read().decode()
        conn.close()
        
        return json.loads(data) if data else {"result": "OK"}
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        return None

def update_status(status="idle", task="", load=0):
    """更新状态"""
    timestamp = int(time.time())
    
    data = {
        "status": status,
        "task": task,
        "load": load,
        "timestamp": timestamp
    }
    
    # 设置当前状态
    redis_request("POST", "/set/kimi:current", {
        "value": json.dumps(data)
    })
    
    # 设置过期时间 5分钟
    redis_request("GET", "/expire/kimi:current/300")
    
    # 添加到历史记录
    redis_request("POST", "/zadd/kimi:history", {
        "score": timestamp,
        "member": json.dumps(data)
    })
    
    print(f"✓ {status} | {task} | {load}%")
    return True

if __name__ == "__main__":
    action = sys.argv[1] if len(sys.argv) > 1 else "heartbeat"
    
    if action == "heartbeat":
        update_status("idle", "等待中", 0)
    elif action == "start":
        task = sys.argv[2] if len(sys.argv) > 2 else "处理中"
        update_status("processing", task, 35)
    elif action == "busy":
        task = sys.argv[2] if len(sys.argv) > 2 else "高负载任务"
        update_status("busy", task, 75)
    elif action == "complete":
        update_status("idle", "已完成", 0)
    else:
        update_status("processing", action, 35)

#!/usr/bin/env python3
"""
Kimi-Claw Status Updater
更新本地状态文件，支持心跳保活和状态变更
"""

import sys
import json
import time
from datetime import datetime, timezone

STATUS_FILE = "/root/.openclaw/workspace/kimi-claw-status/data/current_status.json"

def load_status():
    try:
        with open(STATUS_FILE, 'r') as f:
            return json.load(f)
    except:
        return {"status": "idle", "task": "初始化", "load": 0, "timestamp": 0}

def save_status(status_data):
    with open(STATUS_FILE, 'w') as f:
        json.dump(status_data, f, indent=2)

def get_current_time():
    return int(time.time())

def heartbeat():
    """保持心跳，更新 timestamp"""
    status = load_status()
    status["timestamp"] = get_current_time()
    # 如果超过5分钟没有更新，自动回到 idle
    if status.get("status") == "processing":
        last_update = status.get("timestamp", 0)
        if get_current_time() - last_update > 300:
            status["status"] = "idle"
            status["task"] = "等待指令"
    save_status(status)

def start_task(task_summary):
    """开始新任务"""
    status = {
        "status": "processing",
        "task": task_summary,
        "load": 50,
        "timestamp": get_current_time()
    }
    save_status(status)

def complete_task():
    """完成任务，回到 idle"""
    status = {
        "status": "idle",
        "task": "等待指令",
        "load": 0,
        "timestamp": get_current_time()
    }
    save_status(status)

def main():
    if len(sys.argv) < 2:
        heartbeat()
        return
    
    command = sys.argv[1]
    
    if command == "heartbeat":
        heartbeat()
    elif command == "start" and len(sys.argv) >= 3:
        start_task(sys.argv[2])
    elif command == "complete":
        complete_task()
    else:
        heartbeat()

if __name__ == "__main__":
    main()

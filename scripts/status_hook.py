#!/usr/bin/env python3
"""
对话状态Hook - 在每次对话结束时调用
自动检测任务类型并更新状态
"""

import sys
import subprocess

SCRIPT_PATH = "/root/.openclaw/workspace/kimi-claw-status/scripts/status_updater.py"

def update_after_response(task_summary="", is_complete=False):
    """
    每次我响应后调用这个函数
    
    task_summary: 简短描述当前在做什么（如"UI设计"、"数据分析"）
    is_complete: 是否任务已完成
    """
    if is_complete:
        cmd = ["python3", SCRIPT_PATH, "complete"]
    elif task_summary:
        cmd = ["python3", SCRIPT_PATH, "start", task_summary]
    else:
        cmd = ["python3", SCRIPT_PATH, "heartbeat"]
    
    try:
        subprocess.run(cmd, capture_output=True, timeout=5)
    except Exception:
        pass  # 静默失败，不影响主流程

if __name__ == "__main__":
    # 命令行用法
    if len(sys.argv) > 1:
        task = sys.argv[1]
        is_done = sys.argv[2] == "done" if len(sys.argv) > 2 else False
        update_after_response(task, is_done)
    else:
        update_after_response()

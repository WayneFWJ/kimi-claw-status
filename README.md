# Kimi-Claw Status Monitor

Real-time status monitoring for Kimi-Claw AI assistant using Upstash Redis.

## Monitor Page

https://waynefwj.github.io/kimi-claw-status/realtime-monitor/

## Features

- Real-time status updates via Upstash Redis
- 3-minute heartbeat keepalive
- Automatic status persistence
- Web dashboard for live monitoring

## Status Types

- `idle` - Waiting for tasks
- `processing` - Working on a task
- `busy` - High load task

## Files

- `scripts/status_updater.py` - Status update script
- `realtime-monitor/index.html` - Web dashboard
- `data/current_status.json` - Local status cache

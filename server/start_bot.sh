#!/bin/bash
export PATH="$HOME/.local/bin:/usr/bin:/usr/local/bin:$PATH"
cd ~/yc-voice-agents
nohup uv run bot-gpt.py > /tmp/bot.log 2>&1 &
echo $!

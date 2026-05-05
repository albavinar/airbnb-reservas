#!/bin/bash
# Arranca el Agente SWAPinn en segundo plano y guarda logs
LOG="$HOME/Library/Logs/swapinn-bot.log"
PID_FILE="$HOME/Library/Logs/swapinn-bot.pid"

if [ -f "$PID_FILE" ] && kill -0 "$(cat "$PID_FILE")" 2>/dev/null; then
    echo "El bot ya está en marcha (PID $(cat "$PID_FILE"))."
    exit 0
fi

python3 /Users/albava/scripts/reservas/telegram_bot.py >> "$LOG" 2>&1 &
echo $! > "$PID_FILE"
echo "✅ Agente SWAPinn arrancado (PID $!). Log: $LOG"

#!/bin/bash
echo "=== $(date '+%Y-%m-%d %H:%M:%S') — inicio ==="
cd ~/scripts/reservas
python3 actualizar_reservas.py
echo "=== $(date '+%Y-%m-%d %H:%M:%S') — fin ==="

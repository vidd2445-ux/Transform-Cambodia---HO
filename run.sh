#!/bin/bash
# Transform Cambodia Attendance System - Production Run Script

# Load environment variables from .env
export $(grep -v '^#' .env | grep -v '^$' | xargs)

# Run with gunicorn on all interfaces, port 5000
nohup gunicorn --bind 0.0.0.0:5000 --workers 4 --timeout 120 app:app > access.log 2>&1 &

echo "Server started on 0.0.0.0:5000"
echo "PID: $!"
echo "Logs: access.log"

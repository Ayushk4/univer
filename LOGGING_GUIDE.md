# Structured Logging Guide

This application uses structured logging for audit trails and debugging complex issues.

## 📁 Log Files

All logs are stored in the `logs/` directory:

```
logs/
├── app.log           # Application logs (INFO and WARNING only), JSON format
├── error.log         # Errors only (ERROR and CRITICAL), JSON format
├── audit.log         # Audit trail (user actions), JSON format
└── archived/         # Auto-rotated old logs
    ├── app.log.2025-01-15
    └── audit.log.2025-01-14
```

## 🔄 Log Rotation

- **`app.log`**: INFO and WARNING only, rotates daily at midnight, keeps 30 days
- **`error.log`**: ERROR and CRITICAL only, rotates when it reaches 10MB, keeps 10 files
- **`audit.log`**: All user actions, rotates daily at midnight, keeps 365 days (1 year)

## 📊 Log Formats

### Console Output (Human-Readable)
```
10:30:45 - mcp_server - INFO - Connected to Univer Sheets successfully!
10:31:02 - app - INFO - Query received: Rename sheet0001 to MySheet
```

### JSON Logs (Machine-Readable)

**`app.log`**:
```json
{
  "timestamp": "2025-01-15T10:30:45.123Z",
  "level": "INFO",
  "logger": "mcp_server",
  "message": "Connected to Univer Sheets successfully!",
  "module": "mcp_server",
  "function": "start",
  "line": 91
}
```

**`audit.log`**:
```json
{
  "timestamp": "2025-01-15T10:31:05.234Z",
  "action": "rename_sheet",
  "operation": "rename_sheet",
  "target": "sheet0001",
  "result": "success",
  "duration_ms": 145.67,
  "details": "Renamed 1 sheet(s)"
}
```

## 🔍 Analyzing Logs

### Using `grep`

```bash
# Find all failed operations
grep '"result": "failed"' logs/audit.log

# Find errors from today
grep "$(date +%Y-%m-%d)" logs/error.log

# Track specific sheet modifications
grep '"sheet_name": "MySheet"' logs/audit.log
```

### Using `jq` (Recommended)

```bash
# Find operations that took > 1 second
jq 'select(.duration_ms > 1000)' logs/audit.log

# Count operations by type
jq '.operation' logs/audit.log | sort | uniq -c

# Get all errors from last hour
jq --arg time "$(date -u -d '1 hour ago' +%Y-%m-%dT%H:%M:%S)" \
   'select(.timestamp > $time and .level == "ERROR")' logs/app.log

# Extract failed query details
jq 'select(.action == "query" and .result == "failed")' logs/audit.log
```

### Using Python

```python
import json
from datetime import datetime, timedelta

# Read audit log
with open('logs/audit.log', 'r') as f:
    logs = [json.loads(line) for line in f]

# Find slow operations
slow_ops = [log for log in logs if log.get('duration_ms', 0) > 1000]
print(f"Found {len(slow_ops)} slow operations")

# Get operation statistics
from collections import Counter
ops = Counter(log['operation'] for log in logs)
print(ops.most_common(10))
```

## 📝 Tracked Operations

The following operations are automatically logged with timing and results:

### User Actions (in `audit.log`)
- `query` - AI queries
- `write_data` - Cell data modifications
- `create_sheet` - Sheet creation
- `delete_sheet` - Sheet deletion
- `rename_sheet` - Sheet renaming
- `websocket_connect` - WebSocket connections
- `websocket_disconnect` - WebSocket disconnections

### Fields in Audit Logs
- `timestamp` - ISO 8601 timestamp (UTC)
- `action` - Type of user action
- `operation` - Specific operation performed
- `target` - Affected resource (sheet name, cell range)
- `result` - `success` or `failed`
- `duration_ms` - Operation duration in milliseconds
- `details` - Human-readable description

## 🚨 Monitoring

### Real-Time Monitoring

```bash
# Watch all logs
tail -f logs/app.log | jq '.'

# Watch errors only
tail -f logs/error.log | jq '.level + " - " + .message'

# Watch audit trail
tail -f logs/audit.log | jq '.action + ": " + .details'
```

### Daily Summary Script

```bash
#!/bin/bash
# daily_summary.sh

echo "=== Daily Log Summary ==="
echo "Date: $(date)"
echo

echo "Total Queries:"
grep '"action": "query"' logs/audit.log | wc -l

echo "Failed Operations:"
grep '"result": "failed"' logs/audit.log | wc -l

echo "Average Query Duration (ms):"
jq -r 'select(.action == "query" and .duration_ms != null) | .duration_ms' \
   logs/audit.log | awk '{sum+=$1; count++} END {print sum/count}'

echo "Top Operations:"
jq -r '.operation' logs/audit.log | sort | uniq -c | sort -rn | head -5
```

## 🔐 Security & Privacy

- Logs contain user actions and data modifications
- **Do not commit logs to version control** (already in `.gitignore`)
- Rotate and archive old logs regularly
- Consider encrypting archived logs for long-term storage

## 🛠️ Configuration

To change log levels or rotation settings, edit `logging_config.py`:

```python
# Change console log level
setup_logging(log_level=logging.DEBUG)  # Shows more detail

# Change retention periods
backupCount=30   # Days to keep app.log
backupCount=365  # Days to keep audit.log
```

## 📈 Performance Impact

- Structured logging adds ~2-5ms per operation
- Minimal disk I/O (async writes)
- Log rotation happens in background threads
- No noticeable impact on user experience

## 🐛 Troubleshooting

**Problem**: No logs being created
```bash
# Check if logs directory exists
ls -la logs/

# Check permissions
chmod 755 logs/

# Check if logging_config.py is imported
python -c "import logging_config; print('OK')"
```

**Problem**: Logs growing too large
```bash
# Manually rotate logs
find logs/ -name "*.log" -type f -mtime +30 -delete

# Or adjust rotation settings in logging_config.py
```

**Problem**: Can't parse JSON logs
```bash
# Validate JSON
jq '.' logs/app.log > /dev/null && echo "Valid JSON" || echo "Invalid JSON"

# Find invalid lines
grep -n -v '^{' logs/app.log
```

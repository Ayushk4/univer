"""
Structured logging configuration with audit trail support
"""
import logging
import logging.handlers
import json
import os
from datetime import datetime
from pathlib import Path

# Create logs directory structure
LOG_DIR = Path(__file__).parent / "logs"
LOG_DIR.mkdir(exist_ok=True)
ARCHIVE_DIR = LOG_DIR / "archived"
ARCHIVE_DIR.mkdir(exist_ok=True)


class StructuredFormatter(logging.Formatter):
    """JSON formatter for structured logging"""
    
    def format(self, record):
        log_data = {
            "timestamp": datetime.utcnow().isoformat() + "Z",
            "level": record.levelname,
            "logger": record.name,
            "message": record.getMessage(),
            "module": record.module,
            "function": record.funcName,
            "line": record.lineno,
        }
        
        # Add exception info if present
        if record.exc_info:
            log_data["exception"] = self.formatException(record.exc_info)
        
        # Add custom fields from extra parameter
        if hasattr(record, 'user_action'):
            log_data["user_action"] = record.user_action
        if hasattr(record, 'sheet_name'):
            log_data["sheet_name"] = record.sheet_name
        if hasattr(record, 'operation'):
            log_data["operation"] = record.operation
        if hasattr(record, 'duration_ms'):
            log_data["duration_ms"] = record.duration_ms
        if hasattr(record, 'cell_range'):
            log_data["cell_range"] = record.cell_range
        if hasattr(record, 'result'):
            log_data["result"] = record.result
            
        return json.dumps(log_data)


class AuditFormatter(logging.Formatter):
    """Specialized formatter for audit logs"""
    
    def format(self, record):
        audit_data = {
            "timestamp": datetime.utcnow().isoformat() + "Z",
            "action": getattr(record, 'user_action', 'unknown'),
            "operation": getattr(record, 'operation', ''),
            "target": getattr(record, 'sheet_name', '') or getattr(record, 'cell_range', ''),
            "result": getattr(record, 'result', 'success'),
            "duration_ms": getattr(record, 'duration_ms', None),
            "details": record.getMessage(),
        }
        return json.dumps(audit_data)


class InfoOnlyFilter(logging.Filter):
    """Filter that only allows INFO and WARNING levels (excludes ERROR and above)"""
    
    def filter(self, record):
        # Only allow INFO and WARNING, exclude ERROR and CRITICAL
        return record.levelno < logging.ERROR


def setup_logging(log_level=logging.INFO):
    """Configure logging for the entire application"""
    
    # Root logger configuration
    root_logger = logging.getLogger()
    root_logger.setLevel(logging.DEBUG)  # Capture everything, filter at handler level
    
    # Remove existing handlers to avoid duplicates
    root_logger.handlers = []
    
    # 1. Console Handler (human-readable, with emojis)
    console_handler = logging.StreamHandler()
    console_handler.setLevel(log_level)
    console_formatter = logging.Formatter(
        '%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        datefmt='%H:%M:%S'
    )
    console_handler.setFormatter(console_formatter)
    root_logger.addHandler(console_handler)
    
    # 2. Main App Log (JSON, rotating daily) - INFO and WARNING only
    app_handler = logging.handlers.TimedRotatingFileHandler(
        LOG_DIR / "app.log",
        when="midnight",
        interval=1,
        backupCount=30,  # Keep 30 days
        encoding='utf-8'
    )
    app_handler.setLevel(logging.INFO)
    app_handler.setFormatter(StructuredFormatter())
    app_handler.addFilter(InfoOnlyFilter())  # Exclude ERROR and above
    app_handler.suffix = "%Y-%m-%d"  # Date suffix for archived files
    root_logger.addHandler(app_handler)
    
    # 3. Error Log (JSON, rotating by size)
    error_handler = logging.handlers.RotatingFileHandler(
        LOG_DIR / "error.log",
        maxBytes=10*1024*1024,  # 10MB
        backupCount=10,
        encoding='utf-8'
    )
    error_handler.setLevel(logging.ERROR)
    error_handler.setFormatter(StructuredFormatter())
    root_logger.addHandler(error_handler)
    
    # 4. Audit Log (separate logger, JSON, keep for 1 year)
    audit_logger = logging.getLogger('audit')
    audit_logger.setLevel(logging.INFO)
    audit_logger.propagate = False  # Don't send to root logger
    
    audit_handler = logging.handlers.TimedRotatingFileHandler(
        LOG_DIR / "audit.log",
        when="midnight",
        interval=1,
        backupCount=365,  # Keep 1 year for audit trail
        encoding='utf-8'
    )
    audit_handler.setLevel(logging.INFO)
    audit_handler.setFormatter(AuditFormatter())
    audit_handler.suffix = "%Y-%m-%d"
    audit_logger.addHandler(audit_handler)
    
    logging.info("🚀 Logging system initialized")
    return audit_logger


# Initialize on import
audit_logger = setup_logging()

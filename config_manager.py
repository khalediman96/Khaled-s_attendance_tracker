"""Configuration management module for attendance tracker."""

import json
import os
from pathlib import Path
from typing import Dict, Any
import logging

logger = logging.getLogger(__name__)

class ConfigManager:
    """Manages application configuration and settings."""
    
    DEFAULT_CONFIG = {
        "pdf_path": "",
        "document_path": "",  # For Word documents
        "document_type": "pdf",  # "pdf" or "word"
        "output_directory": "filled_docs",
        "field_names": {
            "time_in": "time_in",
            "time_out": "time_out",
            "date": "date",
            "employee_name": "employee_name"
        },
        "selected_month": "",  # Selected month for attendance tracking
        "auto_start": False,
        "minimize_to_tray": True,
        "log_events": True,
        "log_directory": "logs",
        "date_format": "%Y-%m-%d",
        "time_format": "%H:%M:%S",
        "pdf_fallback": {
            "enabled": True,
            "coordinates": {
                "time_in": [100, 500],
                "time_out": [300, 500],
                "date": [100, 550],
                "employee_name": [100, 600],
                "month": [100, 650],
                "Month": [100, 650],
                "month_year": [100, 650],
                "Month/Year": [100, 650]
            }
        }
    }
    
    def __init__(self, config_file: str = None):
        """Initialize configuration manager.
        
        Args:
            config_file: Path to configuration file (optional, defaults to app data dir)
        """
        import os
        self.app_data_dir = Path(os.getenv('APPDATA') or Path.home() / '.attendance_tracker') / 'AttendanceTracker'
        # Only create app_data_dir if/when a config file is actually written
        if config_file is None:
            self.config_file = self.app_data_dir / 'settings.json'
        else:
            self.config_file = Path(config_file)
        # Update default paths to be inside app_data_dir
        self.DEFAULT_CONFIG['output_directory'] = str(self.app_data_dir / 'filled_docs')
        self.DEFAULT_CONFIG['log_directory'] = str(self.app_data_dir / 'logs')
        if self.config_file.exists():
            self.config = self.load_config()
        else:
            self.config = self.DEFAULT_CONFIG.copy()
    
    def load_config(self) -> Dict[str, Any]:
        """Load configuration from file or create default.
        
        Returns:
            Configuration dictionary
        """
        if self.config_file.exists():
            try:
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    # Merge with defaults to ensure all keys exist
                    return {**self.DEFAULT_CONFIG, **config}
            except Exception as e:
                logger.error(f"Error loading config: {e}")
                return self.DEFAULT_CONFIG.copy()
        else:
            return self.DEFAULT_CONFIG.copy()
    
    def save_config(self, config: Dict[str, Any] = None) -> bool:
        """Save configuration to file.
        
        Args:
            config: Configuration to save (uses current if None)
            
        Returns:
            Success status
        """
        if config is None:
            config = self.config
        try:
            self._ensure_directories()
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(config, f, indent=4)
            self.config = config
            return True
        except Exception as e:
            logger.error(f"Error saving config: {e}")
            return False
    
    def get(self, key: str, default: Any = None) -> Any:
        """Get configuration value.
        
        Args:
            key: Configuration key (supports dot notation)
            default: Default value if key not found
            
        Returns:
            Configuration value
        """
        keys = key.split('.')
        value = self.config
        
        for k in keys:
            if isinstance(value, dict) and k in value:
                value = value[k]
            else:
                return default
        
        return value
    
    def set(self, key: str, value: Any) -> bool:
        """Set configuration value.
        
        Args:
            key: Configuration key (supports dot notation)
            value: Value to set
            
        Returns:
            Success status
        """
        keys = key.split('.')
        config = self.config
        
        # Navigate to parent
        for k in keys[:-1]:
            if k not in config:
                config[k] = {}
            config = config[k]
        
        # Set value
        config[keys[-1]] = value
        return self.save_config()
    
    def _ensure_directories(self):
        """Create necessary directories if they don't exist."""
        directories = [
            self.get('output_directory', 'filled_pdfs'),
            self.get('log_directory', 'logs')
        ]
        
    # Only create directories if/when files are actually written to them
    pass
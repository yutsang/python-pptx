"""
Enhanced logging configuration for FDD application
Creates individual log files for each button/action with timestamps
"""

import logging
import os
from datetime import datetime
from pathlib import Path

class ButtonLogger:
    """Logger that creates individual log files for each button/action"""

    def __init__(self, log_dir="logs"):
        self.log_dir = Path(log_dir)
        self.log_dir.mkdir(exist_ok=True)
        self.session_id = datetime.now().strftime("%Y%m%d_%H%M%S")

        # Setup main logger
        self.main_logger = logging.getLogger('fdd_main')
        self.main_logger.setLevel(logging.INFO)

        # Remove any existing handlers
        for handler in self.main_logger.handlers[:]:
            self.main_logger.removeHandler(handler)

        # Create console handler
        console_handler = logging.StreamHandler()
        console_handler.setLevel(logging.INFO)
        console_formatter = logging.Formatter(
            '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
        )
        console_handler.setFormatter(console_formatter)
        self.main_logger.addHandler(console_handler)

    def get_button_logger(self, button_name):
        """Get a logger for a specific button/action"""
        logger_name = f'fdd_{button_name}_{self.session_id}'
        logger = logging.getLogger(logger_name)
        logger.setLevel(logging.DEBUG)

        # Remove existing handlers to avoid duplicates
        for handler in logger.handlers[:]:
            logger.removeHandler(handler)

        # Create file handler for this button
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        log_file = self.log_dir / f"{button_name}_{timestamp}.log"

        file_handler = logging.FileHandler(log_file, encoding='utf-8')
        file_handler.setLevel(logging.DEBUG)

        # Create detailed formatter
        formatter = logging.Formatter(
            '%(asctime)s - %(name)s - %(levelname)s - %(funcName)s:%(lineno)d - %(message)s'
        )
        file_handler.setFormatter(formatter)

        # Create console handler for this button
        console_handler = logging.StreamHandler()
        console_handler.setLevel(logging.INFO)
        console_formatter = logging.Formatter(
            f'[{button_name}] %(asctime)s - %(levelname)s - %(message)s'
        )
        console_handler.setFormatter(console_formatter)

        logger.addHandler(file_handler)
        logger.addHandler(console_handler)

        return logger

    def log_button_action(self, button_name, action, details=None):
        """Log a button action with details"""
        logger = self.get_button_logger(button_name)

        logger.info(f"Button Action: {action}")
        if details:
            if isinstance(details, dict):
                for key, value in details.items():
                    logger.info(f"  {key}: {value}")
            else:
                logger.info(f"Details: {details}")

        return logger

def create_button_logger():
    """Create and return a ButtonLogger instance"""
    return ButtonLogger()

def log_ai_processing(button_name, agent_name, key, system_prompt, user_prompt, response=None, error=None):
    """Log AI processing details for a specific button"""
    logger = create_button_logger().get_button_logger(button_name)

    logger.info(f"AI Processing - Agent: {agent_name}, Key: {key}")

    # Log prompts
    logger.debug(f"System Prompt Length: {len(system_prompt)} chars")
    logger.debug(f"User Prompt Length: {len(user_prompt)} chars")

    # Log system prompt (truncated for readability)
    logger.debug(f"System Prompt: {system_prompt[:500]}{'...' if len(system_prompt) > 500 else ''}")

    # Log user prompt (truncated for readability)
    logger.debug(f"User Prompt: {user_prompt[:500]}{'...' if len(user_prompt) > 500 else ''}")

    if response:
        logger.info(f"AI Response Length: {len(response)} chars")
        logger.debug(f"AI Response: {response[:500]}{'...' if len(response) > 500 else ''}")

    if error:
        logger.error(f"AI Processing Error: {error}")

    return logger

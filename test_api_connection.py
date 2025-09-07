#!/usr/bin/env python3
"""
Simple test script to test API connection and reproduce the error.
"""

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from app.config import AppConfig
from app.services.api_manager import APIManager
from app.services.gemini_api_manager import GeminiAPIManager
from app.utils.logger import setup_logger

def test_api_connection():
    """Test the API connection"""
    # Setup logger
    logger = setup_logger("APITest", level="DEBUG")

    # Load configuration
    config = AppConfig()

    # Get API settings
    api_type = config.get('api_type', 'openai')
    api_key = config.get('api_key', '')
    gemini_api_key = config.get('gemini_api_key', '')
    model = config.get('model', 'gpt-3.5-turbo')
    gemini_model = config.get('gemini_model', 'gemini-1.5-flash')
    ollama_url = config.get('ollama_url', 'http://localhost:11434')

    logger.info(f"Testing API connection for type: {api_type}")
    logger.info(f"Model: {model if api_type == 'openai' else gemini_model if api_type == 'gemini' else 'ollama'}")
    logger.info(f"Gemini API key present: {bool(gemini_api_key)}")

    # Test Gemini directly if that's the API type
    if api_type == 'gemini':
        logger.info("Testing Gemini API directly...")
        gemini_manager = GeminiAPIManager(api_key=gemini_api_key, model=gemini_model)
        success, message = gemini_manager.test_connection()
        logger.info(f"Direct Gemini test result: {'SUCCESS' if success else 'FAILED'}")
        logger.info(f"Message: {message}")
        return success, message

    # Create API manager
    if api_type == 'gemini':
        api_key = gemini_api_key

    api_manager = APIManager(
        api_key=api_key,
        model=model if api_type == 'openai' else gemini_model,
        api_type=api_type,
        ollama_url=ollama_url
    )

    # Test connection
    logger.info("Starting connection test...")
    success, message = api_manager.test_connection()

    logger.info(f"Connection test result: {'SUCCESS' if success else 'FAILED'}")
    logger.info(f"Message: {message}")

    return success, message

if __name__ == "__main__":
    test_api_connection()
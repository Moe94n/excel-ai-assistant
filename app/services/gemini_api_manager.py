"""
Gemini API Manager for the Excel AI Assistant.
Handles interactions with Google's Gemini AI models.
"""

import logging
import time
from typing import Tuple, Optional, Dict, Any

import google.generativeai as genai
from google.api_core.exceptions import GoogleAPIError


class GeminiAPIManager:
    """Manager for interacting with Google's Gemini AI models"""

    def __init__(self, api_key: str = "", model: str = "gemini-1.5-flash"):
        """Initialize the Gemini API manager"""
        self.api_key = api_key
        self.model = model
        self.client = None
        self.logger = logging.getLogger("GeminiAPIManager")

        # Rate limiting
        self.request_count = 0
        self.request_start_time = time.time()
        self.max_requests_per_minute = 60  # Higher for Gemini

        # Initialize if API key is provided
        if api_key:
            self.initialize()

    def initialize(self, api_key: Optional[str] = None) -> bool:
        """Initialize the Gemini API client"""
        if api_key:
            self.api_key = api_key

        if not self.api_key:
            self.logger.error("API Key is required for Gemini")
            return False

        try:
            genai.configure(api_key=self.api_key)
            self.client = genai.GenerativeModel(self.model)
            return True
        except Exception as e:
            self.logger.error(f"Failed to initialize Gemini client: {e}")
            return False

    def set_api_key(self, api_key: str) -> None:
        """Set the API key"""
        self.api_key = api_key
        self.initialize()

    def set_model(self, model: str) -> None:
        """Set the model to use for API calls"""
        self.model = model
        if self.client:
            self.client = genai.GenerativeModel(model)

    def get_available_models(self) -> Tuple[bool, Optional[list], Optional[str]]:
        """
        Get list of available Gemini models

        Returns:
            Tuple of (success, models_list, error_message)
        """
        try:
            models = genai.list_models()
            # Filter for generative models
            generative_models = [
                {"id": model.name.split("/")[-1], "name": model.display_name, "api": "gemini"}
                for model in models
                if model.name.startswith("models/gemini") or "gemini" in model.name.lower()
            ]
            return True, generative_models, None
        except Exception as e:
            self.logger.error(f"Error listing Gemini models: {str(e)}")
            return False, None, f"Connection error: {str(e)}"

    def test_connection(self) -> Tuple[bool, str]:
        """Test the Gemini API connection"""
        if not self.client:
            if not self.initialize():
                return False, "API client not initialized"

        try:
            self.logger.info("Starting connection test with generate_content call")
            # Simple test generation - NOTE: max_output_tokens should be in generation_config
            generation_config = genai.GenerationConfig(max_output_tokens=10)
            # Updated safety settings using the correct enum values
            safety_settings = [
                {
                    "category": "HARM_CATEGORY_HARASSMENT",
                    "threshold": "BLOCK_ONLY_HIGH",
                },
                {
                    "category": "HARM_CATEGORY_HATE_SPEECH",
                    "threshold": "BLOCK_ONLY_HIGH",
                },
                {
                    "category": "HARM_CATEGORY_SEXUALLY_EXPLICIT",
                    "threshold": "BLOCK_ONLY_HIGH",
                },
                {
                    "category": "HARM_CATEGORY_DANGEROUS_CONTENT",
                    "threshold": "BLOCK_ONLY_HIGH",
                },
            ]
            response = self.client.generate_content(
                "Respond with the word 'success' to confirm API access:",
                generation_config=generation_config,
                safety_settings=safety_settings
            )
            self.logger.info("Connection test completed successfully")
            self.logger.debug(f"Response type: {type(response)}")
            self.logger.debug(f"Response object: {response}")
            if hasattr(response, 'candidates') and response.candidates:
                candidate = response.candidates[0]
                self.logger.debug(f"Candidate finish_reason: {candidate.finish_reason}")
                self.logger.debug(f"Candidate content: {getattr(candidate, 'content', 'No content')}")
            try:
                text_value = response.text
                self.logger.debug(f"Response text: {text_value}")
                # Check if we got a valid response
                if text_value and len(text_value.strip()) > 0:
                    return True, "Connection successful"
                else:
                    return False, "Empty response from API"
            except Exception as e:
                self.logger.error(f"Error accessing response.text: {e}")
                if response and hasattr(response, 'candidates') and response.candidates:
                    candidate = response.candidates[0]
                    if hasattr(candidate, 'finish_reason') and candidate.finish_reason == 2:
                        self.logger.info("Detected finish_reason=2, handling as filtered response")
                        return False, "API blocked response due to safety filters. Please try again in a few moments or try a different request."
                raise e  # Re-raise to handle in outer catch
        except Exception as e:
            self.logger.error(f"Gemini connection test error: {str(e)}")
            # Check if there's a GenerateContentResponse attribute
            if hasattr(e, 'response') and e.response:
                try:
                    resp = e.response
                    self.logger.error(f"Error response candidates: {resp.candidates}")
                    if resp.candidates:
                        cand = resp.candidates[0]
                        self.logger.error(f"Error candidate finish_reason: {cand.finish_reason}")
                        if hasattr(cand, 'finish_reason') and cand.finish_reason == 2:
                            return False, "API blocked response due to safety filters. Please wait a moment and try again."
                except:
                    pass
            return False, f"Error: {str(e)}"
        return True, "Connection successful"

    def process_single_cell(
            self,
            cell_content: str,
            system_prompt: str,
            user_prompt: str,
            temperature: float = 0.3,
            max_tokens: int = 150,
            context_data: Optional[Dict[str, Any]] = None
    ) -> Tuple[bool, Optional[str], Optional[str]]:
        """
        Process a single cell using Gemini

        Args:
            cell_content: The content of the cell being processed
            system_prompt: System prompt for the AI
            user_prompt: User prompt for the AI
            temperature: Temperature parameter for generation
            max_tokens: Maximum tokens in response
            context_data: Optional dictionary with context data from other columns

        Returns:
            Tuple of (success, result, error_message)
        """
        if not self.client:
            if not self.initialize():
                return False, None, "API client not initialized"

        if not self._check_rate_limit():
            return False, None, "Rate limit exceeded. Please try again later."

        formatted_prompt = f"{user_prompt}\n\nCell content: {cell_content}"

        # Add context information if available
        if context_data and len(context_data) > 0:
            context_text = "\n\nContext information:\n"
            for key, value in context_data.items():
                context_text += f"- {key}: {value}\n"
            formatted_prompt += context_text

        try:
            # Configure generation parameters
            generation_config = genai.GenerationConfig(
                temperature=temperature,
                max_output_tokens=max_tokens,
            )
            # Updated safety settings using the correct enum values
            safety_settings = [
                {
                    "category": "HARM_CATEGORY_HARASSMENT",
                    "threshold": "BLOCK_NONE",
                },
                {
                    "category": "HARM_CATEGORY_HATE_SPEECH",
                    "threshold": "BLOCK_NONE",
                },
                {
                    "category": "HARM_CATEGORY_SEXUALLY_EXPLICIT",
                    "threshold": "BLOCK_NONE",
                },
                {
                    "category": "HARM_CATEGORY_DANGEROUS_CONTENT",
                    "threshold": "BLOCK_NONE",
                },
            ]

            # Make API call
            response = self.client.generate_content(
                f"System: {system_prompt}\n\nUser: {formatted_prompt}",
                generation_config=generation_config,
                safety_settings=safety_settings
            )

            # Increment request counter
            self._increment_request_count()

            # Extract result text
            # First check if response was blocked or has issues
            if response and hasattr(response, 'candidates') and response.candidates:
                candidate = response.candidates[0]
                if hasattr(candidate, 'finish_reason') and candidate.finish_reason == 2:
                    return False, None, "Response blocked by safety filters. Try rephrasing your prompt or wait a moment before retrying."
            try:
                text_value = response.text
                if text_value:
                    return True, text_value.strip(), None
                else:
                    return False, None, "Empty response from API"
            except Exception as e:
                self.logger.error(f"Error accessing response.text: {e}")
                return False, None, f"Invalid response: {e}"

        except GoogleAPIError as e:
            error_msg = f"Gemini API Error: {str(e)}"
            self.logger.error(error_msg)
            return False, None, error_msg
        except Exception as e:
            error_msg = f"Error: {str(e)}"
            self.logger.error(error_msg)
            return False, None, error_msg

    def _increment_request_count(self) -> None:
        """Track API request count for rate limiting"""
        current_time = time.time()
        # Reset counter if a minute has passed
        if current_time - self.request_start_time > 60:
            self.request_count = 0
            self.request_start_time = current_time

        self.request_count += 1

    def _check_rate_limit(self) -> bool:
        """Check if we're under the rate limit"""
        current_time = time.time()
        # Reset counter if a minute has passed
        if current_time - self.request_start_time > 60:
            self.request_count = 0
            self.request_start_time = current_time
            return True

        # Check if we're still under the limit
        return self.request_count < self.max_requests_per_minute

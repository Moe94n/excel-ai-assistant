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
        self.logger.info("Attempting to fetch available Gemini models")
        try:
            self.logger.debug("Calling genai.list_models()")
            models = genai.list_models()
            self.logger.debug(f"Retrieved {len(models)} total models from Gemini API")

            # Filter for generative models
            generative_models = [
                {"id": model.name.split("/")[-1], "name": model.display_name, "api": "gemini"}
                for model in models
                if model.name.startswith("models/gemini") or "gemini" in model.name.lower()
            ]
            self.logger.info(f"Filtered to {len(generative_models)} Gemini models")
            self.logger.debug(f"Available models: {[m['id'] for m in generative_models]}")
            return True, generative_models, None
        except Exception as e:
            self.logger.error(f"Error listing Gemini models: {str(e)}")
            self.logger.error(f"Exception type: {type(e).__name__}")
            import traceback
            self.logger.error(f"Traceback: {traceback.format_exc()}")
            return False, None, f"Connection error: {str(e)}"

    def test_connection(self) -> Tuple[bool, str]:
        """Test the Gemini API connection"""
        self.logger.info("Testing Gemini API connection")
        if not self.client:
            self.logger.debug("Client not initialized, attempting to initialize")
            if not self.initialize():
                self.logger.error("Failed to initialize Gemini client")
                return False, "API client not initialized"

        try:
            self.logger.info("Starting connection test with generate_content call")
            # Simple test generation - NOTE: max_output_tokens should be in generation_config
            generation_config = genai.GenerationConfig(max_output_tokens=10)
            self.logger.debug(f"Generation config: {generation_config}")

            # Use more permissive safety settings for connection test
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
            self.logger.debug(f"Safety settings: {safety_settings}")

            # Try multiple test prompts, starting with the safest ones
            test_prompts = [
                "Say 'OK' to confirm connection.",
                "Hello world",
                "Test",
                "Respond with 'success' to confirm API access:"
            ]

            response = None
            test_prompt = None

            for prompt in test_prompts:
                test_prompt = prompt
                self.logger.debug(f"Trying test prompt: '{test_prompt}'")

                try:
                    self.logger.debug("Making test API call")
                    response = self.client.generate_content(
                        test_prompt,
                        generation_config=generation_config,
                        safety_settings=safety_settings
                    )
                    # If we get here without exception, check if it was blocked
                    if hasattr(response, 'candidates') and response.candidates:
                        candidate = response.candidates[0]
                        if hasattr(candidate, 'finish_reason') and candidate.finish_reason == 2:
                            self.logger.debug(f"Prompt '{test_prompt}' was blocked, trying next prompt")
                            continue
                        else:
                            self.logger.debug(f"Prompt '{test_prompt}' succeeded")
                            break
                    else:
                        self.logger.debug(f"Prompt '{test_prompt}' had no candidates, trying next")
                        continue
                except Exception as e:
                    self.logger.debug(f"Prompt '{test_prompt}' failed with exception: {str(e)}")
                    continue

            # If all prompts failed, use the last response for error handling
            if not response:
                self.logger.error("All test prompts failed")
                return False, "All connection test prompts were blocked or failed"

            self.logger.debug(f"Raw response object: {response}")
            self.logger.debug(f"Response attributes: {dir(response)}")

            self.logger.info("Connection test completed successfully")
            self.logger.debug(f"Response type: {type(response)}")

            # Validate response structure
            if not hasattr(response, 'candidates') or not response.candidates:
                self.logger.warning("Response has no candidates")
                self.logger.debug(f"Response candidates attribute: {getattr(response, 'candidates', 'No candidates attr')}")
                return False, "Invalid response structure from API"

            candidate = response.candidates[0]
            self.logger.debug(f"Candidate attributes: {dir(candidate)}")
            self.logger.debug(f"Candidate finish_reason: {getattr(candidate, 'finish_reason', 'Unknown')}")
            self.logger.debug(f"Candidate content: {getattr(candidate, 'content', 'No content')}")
            if hasattr(candidate, 'content') and candidate.content:
                self.logger.debug(f"Candidate content parts: {candidate.content.parts}")

            # Check for blocked content
            if hasattr(candidate, 'finish_reason') and candidate.finish_reason == 2:
                self.logger.warning("Response blocked by safety filters")
                self.logger.debug(f"Blocked candidate details: finish_reason={candidate.finish_reason}")
                return False, f"API blocked response due to safety filters. The test prompt may be flagged by {self.model}. Try switching to an older model like gemini-1.5-flash."

            # Try to get response text
            try:
                text_value = response.text
                self.logger.debug(f"Response text: '{text_value}'")
                # Check if we got a valid response
                if text_value and len(text_value.strip()) > 0:
                    return True, "Connection successful"
                else:
                    self.logger.warning("Empty response from API")
                    return False, "Empty response from API"
            except ValueError as e:
                self.logger.error(f"Error accessing response.text: {e}")
                # Check if response was blocked
                if hasattr(candidate, 'finish_reason') and candidate.finish_reason == 2:
                    self.logger.info("Detected finish_reason=2, handling as filtered response")
                    return False, f"API blocked response due to safety filters. The test prompt may be flagged by {self.model}. Try switching to an older model like gemini-1.5-flash."
                return False, f"Invalid response format: {e}"

        except Exception as e:
            self.logger.error(f"Gemini connection test error: {str(e)}")
            self.logger.error(f"Exception type: {type(e).__name__}")

            # Handle specific API errors
            if hasattr(e, 'code'):
                if e.code == 403:
                    return False, "API key invalid or insufficient permissions"
                elif e.code == 429:
                    return False, "Rate limit exceeded. Please try again later."
                elif e.code == 500:
                    return False, "Server error. Please try again later."

            # Check if there's a GenerateContentResponse attribute
            if hasattr(e, 'response') and e.response:
                try:
                    resp = e.response
                    self.logger.error(f"Error response type: {type(resp)}")
                    self.logger.error(f"Error response attributes: {dir(resp)}")
                    self.logger.error(f"Error response candidates: {getattr(resp, 'candidates', 'None')}")
                    if hasattr(resp, 'candidates') and resp.candidates:
                        cand = resp.candidates[0]
                        self.logger.error(f"Error candidate attributes: {dir(cand)}")
                        self.logger.error(f"Error candidate finish_reason: {getattr(cand, 'finish_reason', 'Unknown')}")
                        self.logger.error(f"Error candidate content: {getattr(cand, 'content', 'No content')}")
                        if hasattr(cand, 'finish_reason') and cand.finish_reason == 2:
                            self.logger.warning("Error response indicates safety filter block")
                            return False, f"API blocked response due to safety filters. The test prompt may be flagged by {self.model}. Try switching to an older model like gemini-1.5-flash."
                except Exception as inner_e:
                    self.logger.error(f"Error parsing error response: {inner_e}")
                    self.logger.error(f"Inner exception type: {type(inner_e)}")

            return False, f"Connection failed: {str(e)}"

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
            self.logger.debug(f"Making Gemini API call with temperature={temperature}, max_tokens={max_tokens}")

            # Configure generation parameters
            generation_config = genai.GenerationConfig(
                temperature=temperature,
                max_output_tokens=max_tokens,
            )

            # Use more permissive safety settings for processing
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
            self.logger.debug("Sending request to Gemini API")
            response = self.client.generate_content(
                f"System: {system_prompt}\n\nUser: {formatted_prompt}",
                generation_config=generation_config,
                safety_settings=safety_settings
            )

            # Increment request counter
            self._increment_request_count()
            self.logger.debug("Request counter incremented")

            # Validate response structure
            if not response or not hasattr(response, 'candidates') or not response.candidates:
                self.logger.error("Invalid response structure: no candidates")
                return False, None, "Invalid response structure from API"

            candidate = response.candidates[0]
            self.logger.debug(f"Response candidate finish_reason: {getattr(candidate, 'finish_reason', 'Unknown')}")

            # Check for blocked content
            if hasattr(candidate, 'finish_reason') and candidate.finish_reason == 2:
                self.logger.warning("Response blocked by safety filters")
                return False, None, f"Response blocked by safety filters. The prompt may be flagged by {self.model}. Try switching to an older model like gemini-1.5-flash or rephrasing your prompt."

            # Extract result text
            try:
                text_value = response.text
                self.logger.debug(f"Extracted response text (length: {len(text_value) if text_value else 0})")
                if text_value and len(text_value.strip()) > 0:
                    return True, text_value.strip(), None
                else:
                    self.logger.warning("Empty response from API")
                    return False, None, "Empty response from API"
            except ValueError as e:
                self.logger.error(f"Error accessing response.text: {e}")
                # Check if response was blocked
                if hasattr(candidate, 'finish_reason') and candidate.finish_reason == 2:
                    return False, None, f"Response blocked by safety filters. The prompt may be flagged by {self.model}. Try switching to an older model like gemini-1.5-flash or rephrasing your prompt."
                return False, None, f"Invalid response format: {e}"

        except GoogleAPIError as e:
            self.logger.error(f"Gemini API Error: {str(e)}")
            # Handle specific API errors
            if hasattr(e, 'code'):
                if e.code == 403:
                    return False, None, "API key invalid or insufficient permissions"
                elif e.code == 429:
                    return False, None, "Rate limit exceeded. Please try again later."
                elif e.code == 500:
                    return False, None, "Server error. Please try again later."
            return False, None, f"Gemini API Error: {str(e)}"
        except Exception as e:
            self.logger.error(f"Unexpected error in process_single_cell: {str(e)}")
            self.logger.error(f"Exception type: {type(e).__name__}")
            import traceback
            self.logger.error(f"Traceback: {traceback.format_exc()}")
            return False, None, f"Unexpected error: {str(e)}"

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

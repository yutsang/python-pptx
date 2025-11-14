import yaml
import os
import time
from typing import Dict, List, Optional, Any
import httpx
from openai import OpenAI, AzureOpenAI
import logging

class AIHelper:
    """
    Reusable AI helper class supporting multiple agents and models.
    Supports: content generation, value checks, content refinement, and formatting checks.
    """
    
    def __init__(
        self, 
        model_type: str = 'deepseek',
        agent_name: str = 'agent_1',
        language: str = 'Eng',
        use_heuristic: bool = False,
        config_path: Optional[str] = None
    ):
        """
        Initialize AIHelper with specified model and agent configuration.
        
        Args:
            model_type: Type of model ('openai', 'local', 'deepseek')
            agent_name: Name of the agent ('agent_1', 'agent_2', 'agent_3', 'agent_4')
            language: Language for prompts ('Eng' or 'Chi')
            use_heuristic: Whether to use heuristic mode instead of AI
            config_path: Path to config file (optional)
        """
        self.model_type = model_type
        self.agent_name = agent_name
        self.language = language
        self.use_heuristic = use_heuristic
        
        # Load configuration
        self.config_path = config_path or os.path.join(os.path.dirname(__file__), 'config.yml')
        self.config_details = self.load_config().get(self.model_type, {})
        
        # Initialize client only if not using heuristic mode
        if not self.use_heuristic:
            self.validate_config()
            self.client, self.model = self.initialize_client()
        else:
            self.client = None
            self.model = None
            
        # Load prompts configuration
        self.prompts_path = os.path.join(os.path.dirname(__file__), 'prompts.yml')
        
        # Setup logging
        self.logger = logging.getLogger(f'AIHelper.{agent_name}')

    def load_config(self) -> Dict:
        """Load configuration from YAML file."""
        try:
            with open(self.config_path, 'r', encoding='utf-8') as f:
                return yaml.safe_load(f) or {}
        except FileNotFoundError:
            raise FileNotFoundError(f"Config file not found: {self.config_path}")
        except yaml.YAMLError as e:
            raise ValueError(f"Error parsing config YAML: {e}")

    def validate_config(self):
        """Validate required configuration keys for the model type."""
        required_keys = {
            'openai': ['api_key', 'api_base', 'chat_model', 'api_version_completion'],
            'local': ['api_base', 'api_key', 'chat_model'],
            'deepseek': ['api_key', 'api_base', 'chat_model']
        }
        
        if self.model_type not in required_keys:
            raise ValueError(f"Invalid model type: {self.model_type}")
        
        missing_keys = [key for key in required_keys[self.model_type] 
                       if key not in self.config_details]
        
        if missing_keys:
            raise ValueError(f"Missing required keys for {self.model_type}: {missing_keys}")
        
        return True

    def initialize_client(self):
        """Initialize the appropriate AI client based on model type."""
        if self.model_type == 'openai':
            client = AzureOpenAI(
                api_key=self.config_details['api_key'], 
                base_url=self.config_details['api_base'],
                api_version=self.config_details['api_version_completion'],
                http_client=httpx.Client(verify=False)
            )
            model = self.config_details['chat_model']
            
        elif self.model_type == 'local':
            client = OpenAI(
                base_url=self.config_details['api_base'],
                api_key=self.config_details['api_key'],
                http_client=httpx.Client(verify=False)
            )
            model = self.config_details['chat_model']
            
        elif self.model_type == 'deepseek':
            client = OpenAI(
                base_url=self.config_details['api_base'],
                api_key=self.config_details['api_key'],
                http_client=httpx.Client(verify=False)
            )
            model = self.config_details['chat_model']
        else:
            raise ValueError(f"Invalid model type: {self.model_type}")
        
        return client, model
    
    def load_prompts(self, agent_name: Optional[str] = None) -> tuple:
        """
        Load system and user prompts for the specified agent and language.
        
        Args:
            agent_name: Agent name override (optional)
            
        Returns:
            Tuple of (system_prompt, user_prompt_template)
        """
        agent = agent_name or self.agent_name
        
        try:
            with open(self.prompts_path, 'r', encoding='utf-8') as f:
                prompts_data = yaml.safe_load(f) or {}
            
            agent_data = prompts_data.get(agent, {})
            lang_data = agent_data.get(self.language, {})
            
            system_prompt = lang_data.get('system_prompt', '')
            user_prompt_template = lang_data.get('user_prompt', '')
            
            return system_prompt, user_prompt_template
            
        except FileNotFoundError:
            raise FileNotFoundError(f"Prompts file not found: {self.prompts_path}")
        except Exception as e:
            self.logger.error(f"Error loading prompts for {agent}: {e}")
            return '', ''
    
    def get_response(
        self, 
        user_prompt: str, 
        system_prompt: Optional[str] = None,
        temperature: float = 0.7,
        max_tokens: Optional[int] = None
    ) -> Dict[str, Any]:
        """
        Get response from AI model or heuristic.
        
        Args:
            user_prompt: User prompt text
            system_prompt: System prompt (optional, will load from config if not provided)
            temperature: Temperature for response generation
            max_tokens: Maximum tokens in response
            
        Returns:
            Dictionary with response data including content, tokens, and duration
        """
        start_time = time.time()
        
        # Use heuristic mode if enabled
        if self.use_heuristic:
            response_content = self._heuristic_response(user_prompt)
            return {
                'content': response_content,
                'mode': 'heuristic',
                'duration': time.time() - start_time,
                'tokens_used': len(response_content.split())
            }
        
        # Load system prompt if not provided
        if system_prompt is None:
            system_prompt, _ = self.load_prompts()
        
        # Prepare messages
        messages = [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt}
        ]
        
        # Get response based on model type
        try:
            response_method = self.client.chat.completions.create
            
            # Common parameters
            params = {
                'model': self.model,
                'messages': messages,
                'temperature': temperature
            }
            
            if max_tokens:
                params['max_tokens'] = max_tokens
            
            if self.model_type == 'openai':
                response = response_method(**params)
                content = response.choices[0].message.content
                tokens_used = response.usage.total_tokens if hasattr(response, 'usage') else 0
                
            elif self.model_type == 'local':
                params['stream'] = True
                response = response_method(**params)
                
                response_buffer = []
                for chunk in response:
                    for choice in chunk.choices:
                        delta = choice.delta
                        if delta and delta.content:
                            response_buffer.append(delta.content)
                
                content = ''.join(response_buffer)
                tokens_used = len(content.split())
                
            elif self.model_type == 'deepseek':
                response = response_method(**params)
                content = response.choices[0].message.content
                tokens_used = response.usage.total_tokens if hasattr(response, 'usage') else 0
            else:
                raise ValueError(f"Invalid model type: {self.model_type}")
            
            duration = time.time() - start_time
            
            return {
                'content': content,
                'mode': 'ai',
                'model_type': self.model_type,
                'model': self.model,
                'duration': duration,
                'tokens_used': tokens_used,
                'agent_name': self.agent_name,
                'language': self.language
            }
            
        except Exception as e:
            self.logger.error(f"Error getting response: {e}")
            raise
    
    def _heuristic_response(self, user_prompt: str) -> str:
        """
        Generate heuristic response without AI (rule-based).
        
        Args:
            user_prompt: User prompt text
            
        Returns:
            Heuristic response string
        """
        # Simple heuristic logic - can be expanded based on requirements
        return f"[Heuristic mode] Processed prompt with {len(user_prompt)} characters."
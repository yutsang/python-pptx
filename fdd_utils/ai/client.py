from __future__ import annotations


from .config import FDDConfig, normalize_language_code
from .prompts import get_prompt_engine
import os
import time
import math
import re as _re_client
from typing import Dict, List, Optional, Any
import httpx
from openai import OpenAI, AzureOpenAI
import logging

from ..financial_common import package_file_path

_REJECTED_PARAM_RE = _re_client.compile(r"'param':\s*'([a-zA-Z_][a-zA-Z0-9_]*)'")


def _extract_rejected_param(exc: Exception) -> Optional[str]:
    """Best-effort extraction of the param name an OpenAI-style 400 rejected.

    Gateways evolve — every new sampling knob this model doesn't support (so
    far: temperature, max_tokens, top_p) would otherwise need its own
    hardcoded keyword check. Instead, read the SAME machine-readable 'param'
    field OpenAI's error body already gives us (checked first on the SDK's
    parsed .body, falling back to a regex over str(exc) for other raise
    shapes), so a NEW unsupported param self-heals without a code change.
    """
    body = getattr(exc, 'body', None)
    if isinstance(body, dict):
        err = body.get('error')
        if isinstance(err, dict) and err.get('param'):
            return str(err['param'])
    match = _REJECTED_PARAM_RE.search(str(exc))
    return match.group(1) if match else None


class AIClient:
    """
    Reusable AI helper class supporting multiple agents and models.
    Supports: content generation, value checks, content refinement, and formatting checks.
    """
    _logged_fallbacks = set()
    
    def __init__(
        self,
        model_type: str = 'deepseek',
        agent_name: str = 'agent_1',
        language: str = 'Eng',
        use_heuristic: bool = False,
        config_path: Optional[str] = None,
        model_name: Optional[str] = None,
    ):
        """
        Initialize AIClient with specified model and agent configuration.

        Args:
            model_type: Type of model ('openai', 'local', 'deepseek', 'workbench')
            agent_name: Name of the agent ('agent_1', 'agent_2', 'agent_3', 'agent_4')
            language: Language for prompts ('Eng' or 'Chi')
            use_heuristic: Whether to use heuristic mode instead of AI
            config_path: Path to config file (optional)
            model_name: Optional specific model id within the provider (e.g. pick
                GPT-5.4 instead of the provider's configured default GPT-5.5).
                Overrides config_details['chat_model'] after config load.
        """
        self.model_type_requested = model_type
        self.model_name_requested = model_name
        self.agent_name = agent_name
        # Normalize "Chn" -> "Chi" once here: AIClient is built by every pipeline
        # entry point, so this guarantees prompt lookups and styling get the right
        # code regardless of which entry point (pipeline, reprompt, validator) ran.
        language = normalize_language_code(language)
        self.language = language
        self.use_heuristic = use_heuristic

        # Load configuration (may resolve e.g. local -> deepseek if local has no api_base/api_key)
        self.config_path = config_path or package_file_path('config.yml')
        self.config_manager = FDDConfig(
            config_path=self.config_path,
            language=language,
            model_type=model_type,
        )
        self.prompt_engine = get_prompt_engine()
        self.model_type = self.config_manager.model_type
        self.full_config = self.config_manager.config
        self.data_format = self.config_manager.get_default_data_format()
        if self.use_heuristic:
            self.config_details = self.full_config.get(self.model_type, {})
        else:
            self.config_details = self.config_manager.get_model_config()
        if model_name:
            # Copy before mutating — config_details may be a reference into the
            # shared, cached config dict; overriding in place would leak this
            # instance's model choice into every other AIClient using the same
            # provider (e.g. a concurrent thread on a different agent/account).
            self.config_details = dict(self.config_details)
            self.config_details['chat_model'] = model_name

        agent_config = self.config_manager.get_agent_config(agent_name)
        self.temperature = agent_config.get('temperature')
        self.max_tokens = agent_config.get('max_tokens')
        self.top_p = agent_config.get('top_p')
        self.frequency_penalty = agent_config.get('frequency_penalty')
        self.presence_penalty = agent_config.get('presence_penalty')
        
        # Initialize client only if not using heuristic mode
        if not self.use_heuristic:
            self.validate_config()
            self.client, self.model = self.initialize_client()
        else:
            self.client = None
            self.model = None
            
        # Setup logging
        self.logger = logging.getLogger(f'AIClient.{agent_name}')
        self._configure_external_logging()
        if self.model_type_requested != self.model_type:
            fallback_key = (agent_name, self.model_type_requested, self.model_type)
            if fallback_key not in self._logged_fallbacks:
                self.logger.debug(
                    "Requested model '%s' is not configured; using '%s' from config.",
                    self.model_type_requested,
                    self.model_type,
                )
                self._logged_fallbacks.add(fallback_key)
        self.logger.debug(f"Initialized {agent_name} with temperature={self.temperature}, max_tokens={self.max_tokens}")


    def validate_config(self):
        """Validate required configuration keys for the model type."""
        self.config_manager.get_model_config()
        return True

    def _configure_external_logging(self):
        """Reduce noisy client logs unless explicitly enabled."""
        suppress_logs = self.config_manager.get_logging_config().get('suppress_http_logs', True)
        debug_enabled = os.getenv('HR_DEBUG') == '1' or os.getenv('FDD_DEBUG') == '1'
        if suppress_logs and not debug_enabled:
            logging.getLogger('httpx').setLevel(logging.WARNING)
            logging.getLogger('openai').setLevel(logging.WARNING)

    def initialize_client(self):
        """Initialize the appropriate AI client based on model type."""
        if self.model_type == 'openai':
            client = AzureOpenAI(
                api_key=self.config_details['api_key'],
                base_url=self.config_details['api_base'],
                api_version=self.config_details['api_version_completion'],
                http_client=httpx.Client(verify=False, timeout=httpx.Timeout(120.0, connect=10.0))
            )
            model = self.config_details['chat_model']

        elif self.model_type == 'workbench':
            # KPMG Workbench gateway (Azure OpenAI-compatible). The gateway
            # requires the subscription key duplicated as a header (not just
            # api_key) plus enterprise billing/routing headers. charge_code and
            # region_override are configurable per config.yml; defaults match
            # the values in the reference snippet.
            subscription_key = self.config_details['api_key']
            headers = {
                'Ocp-Apim-Subscription-Key': subscription_key,
                'x-kpmg-charge-code': str(self.config_details.get('charge_code') or '0000'),
                'x-kpmg-region-override': str(self.config_details.get('region_override') or 'westeurope'),
            }
            self._workbench_headers = headers
            client = AzureOpenAI(
                api_key=subscription_key,
                base_url=self.config_details['api_base'],
                api_version=self.config_details['api_version'],
                default_headers=headers,
                http_client=httpx.Client(verify=False, timeout=httpx.Timeout(180.0, connect=10.0)),
            )
            model = self.config_details['chat_model']

        elif self.model_type == 'local':
            client = OpenAI(
                base_url=self.config_details['api_base'],
                api_key=self.config_details['api_key'],
                http_client=httpx.Client(verify=False, timeout=httpx.Timeout(120.0, connect=10.0))
            )
            model = self.config_details['chat_model']

        elif self.model_type == 'deepseek':
            client = OpenAI(
                base_url=self.config_details['api_base'],
                api_key=self.config_details['api_key'],
                http_client=httpx.Client(verify=False, timeout=httpx.Timeout(120.0, connect=10.0))
            )
            model = self.config_details['chat_model']
        else:
            raise ValueError(f"Invalid model type: {self.model_type}")
        
        return client, model

    def get_agent_settings(self, agent_name: Optional[str] = None) -> Dict[str, Any]:
        """Get config-backed settings for a specific pipeline agent."""
        return self.config_manager.get_agent_config(agent_name or self.agent_name)
    
    def load_prompts(self, agent_name: Optional[str] = None) -> tuple:
        agent = agent_name or self.agent_name
        try:
            return self.prompt_engine.get_agent_defaults(agent, self.language)
        except Exception as e:
            self.logger.error(f"Error loading prompts for {agent}: {e}")
            return '', ''

    @staticmethod
    def _estimate_text_tokens(text: Optional[str]) -> int:
        normalized = str(text or "").strip()
        if not normalized:
            return 0
        return max(1, math.ceil(len(normalized) / 4))

    def _build_logging_metadata(
        self,
        *,
        user_prompt: str,
        system_prompt: str,
        content: str,
        duration: float,
        mode: str,
        temperature: Optional[float] = None,
        max_tokens: Optional[int] = None,
        top_p: Optional[float] = None,
        frequency_penalty: Optional[float] = None,
        presence_penalty: Optional[float] = None,
        prompt_tokens: Optional[int] = None,
        completion_tokens: Optional[int] = None,
        total_tokens: Optional[int] = None,
    ) -> Dict[str, Any]:
        system_prompt = str(system_prompt or "")
        user_prompt = str(user_prompt or "")
        content = str(content or "")

        estimated_system_prompt_tokens = self._estimate_text_tokens(system_prompt)
        estimated_user_prompt_tokens = self._estimate_text_tokens(user_prompt)
        estimated_prompt_tokens = estimated_system_prompt_tokens + estimated_user_prompt_tokens
        estimated_output_tokens = self._estimate_text_tokens(content)
        estimated_total_tokens = estimated_prompt_tokens + estimated_output_tokens

        resolved_total_tokens = (
            total_tokens
            if total_tokens is not None
            else (
                (prompt_tokens or 0) + (completion_tokens or 0)
                if prompt_tokens is not None or completion_tokens is not None
                else estimated_total_tokens
            )
        )

        if prompt_tokens is not None or completion_tokens is not None or total_tokens is not None:
            token_usage_source = "provider_usage"
        elif mode == "heuristic":
            token_usage_source = "heuristic_estimate"
        else:
            token_usage_source = "estimated"

        return {
            "mode": mode,
            "model_type": self.model_type,
            "model": self.model,
            "provider": self.model_type,
            "agent_name": self.agent_name,
            "language": self.language,
            "duration": duration,
            "temperature": temperature if temperature is not None else self.temperature,
            "max_tokens": max_tokens if max_tokens is not None else self.max_tokens,
            "top_p": top_p if top_p is not None else self.top_p,
            "frequency_penalty": (
                frequency_penalty if frequency_penalty is not None else self.frequency_penalty
            ),
            "presence_penalty": (
                presence_penalty if presence_penalty is not None else self.presence_penalty
            ),
            "system_prompt_length": len(system_prompt),
            "user_prompt_length": len(user_prompt),
            "prompt_length": len(system_prompt) + len(user_prompt),
            "output_length": len(content),
            "prompt_tokens": prompt_tokens,
            "completion_tokens": completion_tokens,
            "total_tokens": total_tokens,
            "estimated_system_prompt_tokens": estimated_system_prompt_tokens,
            "estimated_user_prompt_tokens": estimated_user_prompt_tokens,
            "estimated_prompt_tokens": estimated_prompt_tokens,
            "estimated_output_tokens": estimated_output_tokens,
            "estimated_total_tokens": estimated_total_tokens,
            "expected_max_output_tokens": max_tokens if max_tokens is not None else self.max_tokens,
            "tokens_used": resolved_total_tokens,
            "token_usage_source": token_usage_source,
        }

    @staticmethod
    def _coerce_message_content(content: Any) -> str:
        if content is None:
            return ""
        if isinstance(content, str):
            return content
        if isinstance(content, list):
            parts: list[str] = []
            for item in content:
                if isinstance(item, str):
                    parts.append(item)
                    continue
                if isinstance(item, dict):
                    text_value = item.get("text") or item.get("content")
                    if isinstance(text_value, str):
                        parts.append(text_value)
                    continue
                text_value = getattr(item, "text", None) or getattr(item, "content", None)
                if isinstance(text_value, str):
                    parts.append(text_value)
            return "".join(parts)
        return str(content)

    @classmethod
    def _extract_content_from_choice(cls, choice: Any) -> str:
        if choice is None:
            return ""
        if isinstance(choice, dict):
            message = choice.get("message")
            if isinstance(message, dict):
                return cls._coerce_message_content(message.get("content"))
            if message is not None:
                return cls._coerce_message_content(getattr(message, "content", None))
            delta = choice.get("delta")
            if isinstance(delta, dict):
                return cls._coerce_message_content(delta.get("content"))
            if delta is not None:
                return cls._coerce_message_content(getattr(delta, "content", None))
            return cls._coerce_message_content(choice.get("text") or choice.get("content"))

        message = getattr(choice, "message", None)
        if message is not None:
            return cls._coerce_message_content(getattr(message, "content", None))
        delta = getattr(choice, "delta", None)
        if delta is not None:
            return cls._coerce_message_content(getattr(delta, "content", None))
        return cls._coerce_message_content(getattr(choice, "text", None) or getattr(choice, "content", None))

    @classmethod
    def _extract_response_content(cls, response: Any) -> str:
        if response is None:
            return ""
        if isinstance(response, str):
            return response
        if isinstance(response, dict):
            if isinstance(response.get("content"), str):
                return response["content"]
            choices = response.get("choices")
            if isinstance(choices, list) and choices:
                return cls._extract_content_from_choice(choices[0])
            if isinstance(response.get("output_text"), str):
                return response["output_text"]
        choices = getattr(response, "choices", None)
        if choices:
            return cls._extract_content_from_choice(choices[0])
        output_text = getattr(response, "output_text", None)
        if isinstance(output_text, str):
            return output_text
        return cls._coerce_message_content(response)

    @classmethod
    def _extract_stream_response_content(cls, response: Any) -> str:
        if response is None:
            return ""
        if isinstance(response, str):
            return response
        if isinstance(response, dict):
            return cls._extract_response_content(response)

        response_buffer: list[str] = []
        try:
            iterator = iter(response)
        except TypeError:
            return cls._extract_response_content(response)

        for chunk in iterator:
            if isinstance(chunk, str):
                response_buffer.append(chunk)
                continue
            chunk_choices = getattr(chunk, "choices", None)
            if chunk_choices is None and isinstance(chunk, dict):
                chunk_choices = chunk.get("choices")
            if chunk_choices:
                for choice in chunk_choices:
                    piece = cls._extract_content_from_choice(choice)
                    if piece:
                        response_buffer.append(piece)
                continue
            piece = cls._extract_response_content(chunk)
            if piece:
                response_buffer.append(piece)
        return "".join(response_buffer)
    
    def get_response(
        self,
        user_prompt: str,
        system_prompt: Optional[str] = None,
        temperature: Optional[float] = None,
        max_tokens: Optional[int] = None,
        top_p: Optional[float] = None,
        frequency_penalty: Optional[float] = None,
        presence_penalty: Optional[float] = None,
        allow_thinking: Optional[bool] = None,
        reasoning_effort: Optional[str] = None,
    ) -> Dict[str, Any]:
        """
        Get response from AI model or heuristic.

        Args:
            user_prompt: User prompt text
            system_prompt: System prompt (optional, will load from config if not provided)
            temperature: Temperature for response generation (optional, uses config default)
            max_tokens: Maximum tokens in response (optional, uses config default)
            top_p: Nucleus sampling parameter (optional, uses config default)
            frequency_penalty: Frequency penalty (optional, uses config default)
            presence_penalty: Presence penalty (optional, uses config default)
            reasoning_effort: Per-call override for workbench.reasoning_effort (optional,
                falls back to the provider-level config default — see agents.*.reasoning_effort)

        Returns:
            Dictionary with response data including content, tokens, and duration
        """
        start_time = time.time()

        # Use config defaults if not provided
        temperature = temperature if temperature is not None else self.temperature
        max_tokens = max_tokens if max_tokens is not None else self.max_tokens
        top_p = top_p if top_p is not None else self.top_p
        frequency_penalty = frequency_penalty if frequency_penalty is not None else self.frequency_penalty
        presence_penalty = presence_penalty if presence_penalty is not None else self.presence_penalty
        effective_reasoning_effort = (
            reasoning_effort if reasoning_effort is not None else self.config_details.get('reasoning_effort')
        )
        
        # Use heuristic mode if enabled
        if self.use_heuristic:
            response_content = self._heuristic_response(user_prompt)
            duration = time.time() - start_time
            response = {
                'content': response_content,
                'mode': 'heuristic',
                'duration': duration,
                'temperature': temperature,
                'max_tokens': max_tokens
            }
            response.update(
                self._build_logging_metadata(
                    user_prompt=user_prompt,
                    system_prompt=system_prompt or "",
                    content=response_content,
                    duration=duration,
                    mode='heuristic',
                    temperature=temperature,
                    max_tokens=max_tokens,
                    top_p=top_p,
                    frequency_penalty=frequency_penalty,
                    presence_penalty=presence_penalty,
                )
            )
            response["temperature"] = temperature
            response["max_tokens"] = max_tokens
            return response
        
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
            
            # Add optional parameters if provided
            if max_tokens:
                params['max_tokens'] = max_tokens
            if top_p is not None:
                params['top_p'] = top_p
            if frequency_penalty is not None:
                params['frequency_penalty'] = frequency_penalty
            if presence_penalty is not None:
                params['presence_penalty'] = presence_penalty
            
            if self.model_type == 'openai':
                response = response_method(**params)
                content = self._extract_response_content(response)
                usage = getattr(response, 'usage', None)

            elif self.model_type == 'workbench':
                # KPMG's reference snippet sends the auth/billing headers BOTH as
                # default_headers (set once on the client) and as extra_headers
                # on every call — keep both, since some gateway configurations
                # only honour per-call headers.
                wb_params = dict(params)
                wb_params['extra_headers'] = dict(getattr(self, '_workbench_headers', {}) or {})

                # Model capability flags — config-driven so a new deployment's
                # quirks (confirmed against the real gateway, not guessed) can
                # be tuned in config.yml without a code change. Defaults match
                # the GPT-5-class reasoning models currently behind this
                # gateway, which reject temperature, top_p, frequency_penalty,
                # and presence_penalty outright (reasoning models don't expose
                # traditional sampling controls) and rename max_tokens.
                if not bool(self.config_details.get('supports_temperature', False)):
                    wb_params.pop('temperature', None)
                if not bool(self.config_details.get('supports_sampling_params', False)):
                    for _p in ('top_p', 'frequency_penalty', 'presence_penalty'):
                        wb_params.pop(_p, None)
                if bool(self.config_details.get('use_max_completion_tokens', True)) and 'max_tokens' in wb_params:
                    wb_params['max_completion_tokens'] = wb_params.pop('max_tokens')
                if effective_reasoning_effort:
                    wb_params['reasoning_effort'] = effective_reasoning_effort

                # Reasoning models spend max_completion_tokens on HIDDEN
                # reasoning tokens before the visible answer — the per-agent
                # max_tokens values in config.yml (1200-1400) are sized for
                # local Qwen3, which has no hidden-token cost for structured
                # stages (allow_thinking=false skips <think> entirely). For
                # workbench, a complex multi-component account (e.g.
                # Investment properties, Operating costs) can exhaust that
                # budget on reasoning alone with reasoning_effort=high,
                # leaving nothing for the answer — a real 200/no-exception
                # response with EMPTY content (confirmed via inputs/昆山.xlsx:
                # 2/20 accounts came back blank). min_max_tokens is a floor,
                # not a replacement — it only raises the budget when the
                # agent's own setting is smaller.
                min_max_tokens = self.config_details.get('min_max_tokens')
                if min_max_tokens:
                    current_budget = wb_params.get('max_completion_tokens', wb_params.get('max_tokens', 0)) or 0
                    if current_budget < int(min_max_tokens):
                        budget_key = 'max_completion_tokens' if 'max_completion_tokens' in wb_params else 'max_tokens'
                        wb_params[budget_key] = int(min_max_tokens)

                _MAX_TOKENS_ALIASES = ('max_tokens', 'max_completion_tokens')

                def _drop_rejected_param(exc: Exception) -> bool:
                    """Read the 'param' the gateway's 400 named and adjust
                    wb_params generically — no hardcoded keyword list, so a
                    param we haven't seen yet (the config flags above only
                    cover known ones) still self-heals. Returns True if
                    wb_params changed (worth retrying)."""
                    param = _extract_rejected_param(exc)
                    if not param:
                        return False
                    if param in _MAX_TOKENS_ALIASES:
                        other = 'max_completion_tokens' if param == 'max_tokens' else 'max_tokens'
                        if param in wb_params:
                            wb_params[other] = wb_params.pop(param)
                            return True
                        return False
                    if param in wb_params:
                        wb_params.pop(param, None)
                        return True
                    return False

                try:
                    response = response_method(**wb_params)
                except TypeError:
                    # SDK-level rejection of a kwarg (older openai-python that
                    # doesn't know reasoning_effort at all) — drop and retry.
                    wb_params.pop('reasoning_effort', None)
                    response = response_method(**wb_params)
                except Exception as first_exc:
                    # Gateway-level rejection (400) — retry a few times in case
                    # dropping one param surfaces another unsupported one right
                    # after (seen in practice: temperature, then top_p, then
                    # frequency_penalty, one at a time). Uses an explicit
                    # sentinel-based loop (not for/else + bare raise) so the
                    # re-raised exception is always the exact object we mean.
                    last_exc = first_exc
                    response = None
                    for _ in range(4):
                        if not _drop_rejected_param(last_exc):
                            raise last_exc
                        try:
                            response = response_method(**wb_params)
                            break
                        except Exception as retry_exc:
                            last_exc = retry_exc
                    if response is None:
                        raise last_exc
                content = self._extract_response_content(response)
                usage = getattr(response, 'usage', None)

            elif self.model_type == 'local':
                params['stream'] = True
                # Qwen3: turn OFF thinking for structured/JSON stages (Auditor,
                # Validator) — it adds latency and pollutes output with no quality
                # gain. vLLM/SGLang honour chat_template_kwargs.enable_thinking.
                # If a different server rejects the field, retry without it
                # (strip_thinking still cleans any inline <think> as a fallback).
                if allow_thinking is False:
                    local_params = dict(params)
                    local_params['extra_body'] = {'chat_template_kwargs': {'enable_thinking': False}}
                    try:
                        response = response_method(**local_params)
                    except TypeError:
                        response = response_method(**params)
                    except Exception as exc:
                        if 'extra_body' in str(exc) or 'enable_thinking' in str(exc) or 'chat_template' in str(exc):
                            response = response_method(**params)
                        else:
                            raise
                else:
                    response = response_method(**params)

                content = self._extract_stream_response_content(response)
                usage = getattr(response, 'usage', None)
                
            elif self.model_type == 'deepseek':
                response = response_method(**params)
                content = self._extract_response_content(response)
                usage = getattr(response, 'usage', None)
            else:
                raise ValueError(f"Invalid model type: {self.model_type}")
            
            duration = time.time() - start_time
            prompt_tokens = getattr(usage, 'prompt_tokens', None) if usage is not None else None
            completion_tokens = getattr(usage, 'completion_tokens', None) if usage is not None else None
            total_tokens = getattr(usage, 'total_tokens', None) if usage is not None else None
            
            response_payload = {
                'content': content,
                'mode': 'ai',
                'duration': duration,
                'temperature': temperature,
                'max_tokens': max_tokens,
                'top_p': top_p,
                'frequency_penalty': frequency_penalty,
                'presence_penalty': presence_penalty,
            }
            response_payload.update(
                self._build_logging_metadata(
                    user_prompt=user_prompt,
                    system_prompt=system_prompt,
                    content=content,
                    duration=duration,
                    mode='ai',
                    temperature=temperature,
                    max_tokens=max_tokens,
                    top_p=top_p,
                    frequency_penalty=frequency_penalty,
                    presence_penalty=presence_penalty,
                    prompt_tokens=prompt_tokens,
                    completion_tokens=completion_tokens,
                    total_tokens=total_tokens,
                )
            )
            response_payload["temperature"] = temperature
            response_payload["max_tokens"] = max_tokens
            response_payload["top_p"] = top_p
            response_payload["frequency_penalty"] = frequency_penalty
            response_payload["presence_penalty"] = presence_penalty
            return response_payload
            
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
# --- end ai/client.py ---

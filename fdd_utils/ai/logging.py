from __future__ import annotations

"""
Run logging for the FDD AI pipeline.
"""

import logging
import os
from datetime import datetime
from typing import Any, Dict, Optional

import yaml


class PipelineRunLogger:
    """Unified logger for an FDD AI processing run."""

    def __init__(self, log_dir: str = "fdd_utils/logs", output_dir: str = "fdd_utils/output", debug_mode: bool = False):
        self.log_dir = log_dir
        self.output_dir = output_dir
        self.debug_mode = debug_mode
        os.makedirs(log_dir, exist_ok=True)
        os.makedirs(output_dir, exist_ok=True)

        self.run_id = datetime.now().strftime("%Y%m%d_%H%M%S")
        self.run_folder = os.path.join(log_dir, f"run_{self.run_id}")
        os.makedirs(self.run_folder, exist_ok=True)

        self.log_file = os.path.join(self.run_folder, "processing.log")
        self.log_data_file = os.path.join(self.run_folder, "data.yml")
        self.results_file = os.path.join(self.run_folder, "results.yml")

        self.logger = logging.getLogger(f"ContentGeneration_{self.run_id}")
        self.logger.setLevel(logging.DEBUG if debug_mode else logging.INFO)
        self.logger.handlers = []
        self.logger.propagate = False

        file_handler = logging.FileHandler(self.log_file, encoding="utf-8")
        file_handler.setLevel(logging.DEBUG)

        formatter = logging.Formatter("%(asctime)s - %(name)s - %(levelname)s - %(message)s")
        file_handler.setFormatter(formatter)
        self.logger.addHandler(file_handler)

        if debug_mode:
            self.debug_log_file = os.path.join(self.run_folder, "debug.log")
            debug_handler = logging.FileHandler(self.debug_log_file, encoding="utf-8")
            debug_handler.setLevel(logging.DEBUG)
            debug_handler.setFormatter(formatter)
            self.logger.addHandler(debug_handler)

        self.run_data: Dict[str, Any] = {
            "run_id": self.run_id,
            "start_time": datetime.now().isoformat(),
            "debug_mode": debug_mode,
            "agents_executed": [],
            "processing_results": {},
        }
        self.logger.info("=== Started new AI processing run: %s (debug_mode=%s) ===", self.run_id, debug_mode)

    def _display_name(self, agent_name: str) -> str:
        names = {
            "subagent_1": "Generator",
            "subagent_2": "Auditor",
            "subagent_3": "Refiner",
            "subagent_4": "Validator",
        }
        return names.get(agent_name, agent_name)

    def log_debug(self, category: str, mapping_key: str, message: str, data: Any = None) -> None:
        if not self.debug_mode:
            return
        self.logger.debug("[DEBUG][%s] %s: %s", category, mapping_key, message)
        if data is not None:
            data_str = str(data)
            if len(data_str) > 4000:
                data_str = data_str[:4000] + "... [truncated]"
            self.logger.debug("[DEBUG][%s] %s: DATA:\n%s", category, mapping_key, data_str)

    def log_agent_start(self, agent_name: str, mapping_key: str):
        self.logger.info("[%s] Processing: %s", self._display_name(agent_name), mapping_key)

    def log_agent_complete(
        self,
        agent_name: str,
        mapping_key: str,
        result: Dict[str, Any],
        system_prompt: str = "",
        user_prompt: str = "",
        prompt_context: Optional[Dict[str, Any]] = None,
    ):
        duration = result.get("duration", 0)
        tokens = result.get("tokens_used", 0)
        content = result.get("content", "")
        prompt_length = result.get("prompt_length", len(system_prompt) + len(user_prompt))
        output_length = result.get("output_length", len(content))
        prompt_tokens = result.get("prompt_tokens")
        completion_tokens = result.get("completion_tokens")
        total_tokens = result.get("total_tokens")
        estimated_prompt_tokens = result.get("estimated_prompt_tokens")
        estimated_output_tokens = result.get("estimated_output_tokens")
        expected_max_output_tokens = result.get("expected_max_output_tokens")
        model = result.get("model") or "-"
        model_type = result.get("model_type") or result.get("provider") or result.get("mode", "ai")
        token_usage_source = result.get("token_usage_source", "unknown")

        self.logger.info(
            "[%s] Processed: %s | Duration: %.2fs | Model: %s/%s | Prompt chars: %s | Output chars: %s | Tokens used: %s | Prompt tokens: %s | Completion tokens: %s | Total tokens: %s | Estimated prompt tokens: %s | Estimated output tokens: %s | Expected max output tokens: %s | Token source: %s",
            self._display_name(agent_name),
            mapping_key,
            duration,
            model_type,
            model,
            prompt_length,
            output_length,
            tokens,
            prompt_tokens if prompt_tokens is not None else "-",
            completion_tokens if completion_tokens is not None else "-",
            total_tokens if total_tokens is not None else "-",
            estimated_prompt_tokens if estimated_prompt_tokens is not None else "-",
            estimated_output_tokens if estimated_output_tokens is not None else "-",
            expected_max_output_tokens if expected_max_output_tokens is not None else "-",
            token_usage_source,
        )

        prompt_context = prompt_context or {}
        rhs_summary = prompt_context.get("rhs_remark_summary") or []
        supporting_notes = prompt_context.get("supporting_notes") or []
        table_linked_remarks = prompt_context.get("table_linked_remarks") or []
        user_comment = str(prompt_context.get("user_comment") or "").strip()
        context_fragments = [
            f"supporting_notes={len(supporting_notes)}",
            f"rhs_rows={int(prompt_context.get('rhs_remark_count') or 0)}",
            f"rhs_summary={len(rhs_summary)}",
            f"table_linked_remarks={len(table_linked_remarks)}",
        ]
        if rhs_summary:
            context_fragments.append(
                "rhs_preview=" + " || ".join(str(item).strip() for item in rhs_summary[:3] if str(item).strip())
            )
        if user_comment:
            context_fragments.append(f"user_comment={user_comment[:240]}")
        if prompt_context.get("has_previous_output"):
            context_fragments.append("reprompt_baseline=yes")
        self.logger.info(
            "[%s] Prompt context: %s | %s",
            self._display_name(agent_name),
            mapping_key,
            " | ".join(fragment for fragment in context_fragments if fragment),
        )

        self.run_data["processing_results"].setdefault(mapping_key, {})
        self.run_data["processing_results"][mapping_key][agent_name] = {
            "duration": duration,
            "tokens_used": tokens,
            "mode": result.get("mode", "ai"),
            "model_type": model_type,
            "model": model,
            "provider": result.get("provider", model_type),
            "agent_name": result.get("agent_name", agent_name),
            "language": result.get("language"),
            "temperature": result.get("temperature"),
            "max_tokens": result.get("max_tokens"),
            "top_p": result.get("top_p"),
            "frequency_penalty": result.get("frequency_penalty"),
            "presence_penalty": result.get("presence_penalty"),
            "timestamp": datetime.now().isoformat(),
            "system_prompt": system_prompt or "",
            "user_prompt": user_prompt or "",
            "prompt_context": prompt_context or {},
            "output": content,
            "system_prompt_length": result.get("system_prompt_length", len(system_prompt)),
            "user_prompt_length": result.get("user_prompt_length", len(user_prompt)),
            "prompt_length": prompt_length,
            "content_length": output_length,
            "output_length": output_length,
            "prompt_tokens": prompt_tokens,
            "completion_tokens": completion_tokens,
            "total_tokens": total_tokens,
            "estimated_system_prompt_tokens": result.get("estimated_system_prompt_tokens"),
            "estimated_user_prompt_tokens": result.get("estimated_user_prompt_tokens"),
            "estimated_prompt_tokens": estimated_prompt_tokens,
            "estimated_output_tokens": estimated_output_tokens,
            "estimated_total_tokens": result.get("estimated_total_tokens"),
            "expected_max_output_tokens": expected_max_output_tokens,
            "token_usage_source": token_usage_source,
        }

    def log_error(self, agent_name: str, mapping_key: str, error: Exception):
        self.logger.error("[%s] Error processing %s: %s", self._display_name(agent_name), mapping_key, error)
        self.run_data["processing_results"].setdefault(mapping_key, {})
        self.run_data["processing_results"][mapping_key][agent_name] = {
            "error": str(error),
            "timestamp": datetime.now().isoformat(),
        }

    def finalize(self, results: Optional[Dict[str, Dict[str, str]]] = None):
        self.run_data["end_time"] = datetime.now().isoformat()

        total_duration = 0
        total_tokens = 0
        total_items = len(self.run_data["processing_results"])

        for _key, agents in self.run_data["processing_results"].items():
            for _agent, data in agents.items():
                total_duration += data.get("duration", 0)
                total_tokens += data.get("tokens_used", 0)

        self.run_data["summary"] = {
            "total_items_processed": total_items,
            "total_duration_seconds": total_duration,
            "total_tokens_used": total_tokens,
            "total_prompt_length_chars": sum(
                data.get("prompt_length", 0)
                for agents in self.run_data["processing_results"].values()
                for data in agents.values()
            ),
            "total_output_length_chars": sum(
                data.get("output_length", data.get("content_length", 0))
                for agents in self.run_data["processing_results"].values()
                for data in agents.values()
            ),
            "total_estimated_prompt_tokens": sum(
                data.get("estimated_prompt_tokens", 0)
                for agents in self.run_data["processing_results"].values()
                for data in agents.values()
            ),
            "total_estimated_output_tokens": sum(
                data.get("estimated_output_tokens", 0)
                for agents in self.run_data["processing_results"].values()
                for data in agents.values()
            ),
            "agents_used": list(
                {
                    agent
                    for agents in self.run_data["processing_results"].values()
                    for agent in agents.keys()
                }
            ),
        }

        with open(self.log_data_file, "w", encoding="utf-8") as file:
            yaml.dump(self.run_data, file, default_flow_style=False, allow_unicode=True)

        if results:
            with open(self.results_file, "w", encoding="utf-8") as file:
                yaml.dump(results, file, default_flow_style=False, allow_unicode=True)

        self.logger.info("=== Completed AI processing run: %s ===", self.run_id)
        self.logger.info(
            "Summary: %s items, %.2fs, %s tokens",
            total_items,
            total_duration,
            total_tokens,
        )
# --- end ai/logging.py ---

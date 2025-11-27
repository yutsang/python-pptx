import json
from pathlib import Path
import datetime


class AIAgentLogger:
    """Records each button click and associated AI I/O in a single JSON file under fdd_utils/logs/."""
    def __init__(self):
        from pathlib import Path
        self.log_dir = Path("fdd_utils/logs")
        self.log_dir.mkdir(parents=True, exist_ok=True)
        self.session_id = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        self.session_file = self.log_dir / f"session_{self.session_id}.json"
        self._data = {"session": {"id": self.session_id, "started": datetime.datetime.now().isoformat()}, "events": []}

    def append(self, event: dict):
        try:
            event["ts"] = datetime.datetime.now().isoformat()
            self._data["events"].append(event)
            with open(self.session_file, 'w', encoding='utf-8') as f:
                json.dump(self._data, f, indent=2, ensure_ascii=False)
            print(f"📝 Logged event: {event.get('type', 'unknown')} - {event.get('name', '')} to {self.session_file}")
        except Exception as e:
            print(f"❌ Logging error: {e}")

    # Enhanced wrappers with token counting
    def log_click(self, name: str, payload: dict):
        self.append({"type": "click", "name": name, "payload": payload})
    
    def log_ai_input(self, agent: str, key: str, system: str, user: str, tokens: int = 0):
        self.append({
            "type": "ai_input", 
            "agent": agent, 
            "key": key, 
            "system_prompt": system[:500] + "..." if len(system) > 500 else system,  # Truncate for readability
            "user_prompt": user[:500] + "..." if len(user) > 500 else user,  # Truncate for readability
            "input_tokens": tokens
        })
    
    def log_ai_output(self, agent: str, key: str, output, tokens: int = 0, processing_time: float = 0):
        output_str = str(output)[:500] + "..." if len(str(output)) > 500 else str(output)
        self.append({
            "type": "ai_output", 
            "agent": agent, 
            "key": key, 
            "output": output_str,
            "output_tokens": tokens,
            "processing_time_seconds": processing_time
        })
    
    def log_error(self, agent: str, key: str, error: str):
        self.append({"type": "error", "agent": agent, "key": key, "error": error})
    
    def log_processing_start(self, statement_type: str, keys: list, entity: str):
        self.append({
            "type": "processing_start",
            "statement_type": statement_type,
            "keys_count": len(keys),
            "keys": keys,
            "entity": entity
        })
    
    def log_processing_complete(self, statement_type: str, success: bool, total_time: float, keys_processed: int):
        self.append({
            "type": "processing_complete",
            "statement_type": statement_type,
            "success": success,
            "total_time_seconds": total_time,
            "keys_processed": keys_processed
        })
    
    def log_token_usage(self, key: str, agent: str, input_tokens: int, output_tokens: int, cost: float = 0):
        self.append({
            "type": "token_usage",
            "key": key,
            "agent": agent,
            "input_tokens": input_tokens,
            "output_tokens": output_tokens,
            "total_tokens": input_tokens + output_tokens,
            "estimated_cost_usd": cost
        })


def derive_entity_parts(full_name: str):
    """Return base city and cumulative suffix list from a full entity name.
    Example: 'Haining Wanpu Limited' -> ('Haining', ['Wanpu', 'Wanpu Limited', 'Limited'])
    """
    try:
        tokens = [t for t in str(full_name).strip().split() if t]
        if not tokens:
            return full_name.strip(), []
        base = tokens[0]
        suffixes = []
        for i in range(1, len(tokens)):
            part = " ".join(tokens[1:i+1]).strip()
            if part:
                suffixes.append(part)
        if len(tokens) > 1 and tokens[-1] not in suffixes:
            suffixes.append(tokens[-1])
        return base, suffixes
    except Exception:
        return full_name, []


def get_financial_keys(mapping_path: str = 'fdd_utils/mapping.json'):
    """Get all financial keys from mapping.json or fallback list on error."""
    try:
        with open(mapping_path, 'r') as f:
            mapping = json.load(f)
        return list(mapping.keys())
    except Exception:
        return [
            "Cash", "AR", "Prepayments", "OR", "Other CA", "IP", "Other NCA",
            "AP", "Taxes payable", "OP", "Capital", "Reserve"
        ]


def get_key_display_name(key: str, mapping_path: str = 'fdd_utils/mapping.json'):
    """Best-effort display name for a key using mapping.json; falls back to key."""
    try:
        with open(mapping_path, 'r') as f:
            mapping = json.load(f)

        if key in mapping and mapping[key]:
            values = mapping[key]
            priority_keywords = [
                'Long-term', 'Investment', 'Accounts', 'Other', 'Capital', 'Reserve',
                'Income', 'Expenses', 'Tax', 'Credit', 'Non-operating', 'Advances'
            ]
            for value in values:
                if any(keyword.lower() in value.lower() for keyword in priority_keywords):
                    return value
            for value in values:
                if len(value) > 3 and not value.isupper():
                    return value
            return values[0]
        return key
    except Exception:
        try:
            from fdd_utils.mappings import DISPLAY_NAME_MAPPING_DEFAULT as default_names
            return default_names.get(key, key)
        except Exception:
            return key



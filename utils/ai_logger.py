#!/usr/bin/env python3
"""
AI Logging System for Due Diligence Automation

This module handles logging of all AI agent interactions including:
- Input prompts (system and user)
- AI responses
- Processing metadata
- Error handling
"""

import json
import os
import datetime
from pathlib import Path
from typing import Dict, Any, Optional

class AILogger:
    """Comprehensive AI interaction logging system"""
    
    def __init__(self, base_dir: str = "logging"):
        self.base_dir = Path(base_dir)
        self.base_dir.mkdir(exist_ok=True)
        
        # Create timestamped session directory
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        self.session_dir = self.base_dir / f"ai_session_{timestamp}"
        self.session_dir.mkdir(exist_ok=True)
        
        # Initialize log files
        self.session_log = self.session_dir / "session_summary.json"
        self.detailed_log = self.session_dir / "detailed_interactions.jsonl"
        
        # Session metadata
        self.session_data = {
            "session_id": timestamp,
            "start_time": datetime.datetime.now().isoformat(),
            "interactions": [],
            "summary": {
                "total_interactions": 0,
                "successful_responses": 0,
                "errors": 0,
                "agents_used": set(),
                "keys_processed": []
            }
        }
        
        print(f"🗂️ AI logging session started: {self.session_dir}")
    
    def log_ai_interaction(
        self, 
        agent_name: str, 
        key: str, 
        system_prompt: str, 
        user_prompt: str, 
        ai_response: str,
        entity_name: str = "",
        processing_time: float = 0,
        error: Optional[str] = None,
        ai_connection_status: str = "unknown"
    ):
        """Log a complete AI interaction"""
        
        timestamp = datetime.datetime.now().isoformat()
        interaction_id = f"{agent_name}_{key}_{len(self.session_data['interactions'])}"
        
        interaction_data = {
            "interaction_id": interaction_id,
            "timestamp": timestamp,
            "agent_name": agent_name,
            "key": key,
            "entity_name": entity_name,
            "processing_time_seconds": processing_time,
            "ai_connection_status": ai_connection_status,
            "input": {
                "system_prompt": system_prompt,
                "user_prompt": user_prompt,
                "system_prompt_length": len(system_prompt),
                "user_prompt_length": len(user_prompt)
            },
            "output": {
                "ai_response": ai_response,
                "response_length": len(ai_response),
                "success": error is None
            },
            "error": error,
            "metadata": {
                "is_placeholder": "[Demo AI Analysis]" in ai_response or "[AI Response Placeholder]" in ai_response,
                "is_fallback": "[Fallback Response]" in ai_response,
                "has_real_ai": not any(x in ai_response for x in ["[Demo AI Analysis]", "[AI Response Placeholder]", "[Fallback Response]"]),
                "connection_status": ai_connection_status
            }
        }
        
        # Add to session data
        self.session_data["interactions"].append(interaction_data)
        self.session_data["summary"]["total_interactions"] += 1
        self.session_data["summary"]["agents_used"].add(agent_name)
        
        if key not in self.session_data["summary"]["keys_processed"]:
            self.session_data["summary"]["keys_processed"].append(key)
        
        if error:
            self.session_data["summary"]["errors"] += 1
        else:
            self.session_data["summary"]["successful_responses"] += 1
        
        # Write detailed log entry (JSONL format)
        with open(self.detailed_log, 'a', encoding='utf-8') as f:
            f.write(json.dumps(interaction_data, ensure_ascii=False) + '\n')
        
        # Save individual interaction file
        interaction_file = self.session_dir / f"{interaction_id}.json"
        with open(interaction_file, 'w', encoding='utf-8') as f:
            json.dump(interaction_data, f, indent=2, ensure_ascii=False)
        
        print(f"📝 Logged AI interaction: {agent_name} -> {key}")
        return interaction_id
    
    def finalize_session(self):
        """Finalize the logging session and save summary"""
        self.session_data["end_time"] = datetime.datetime.now().isoformat()
        self.session_data["summary"]["agents_used"] = list(self.session_data["summary"]["agents_used"])
        
        # Calculate session statistics
        duration = datetime.datetime.fromisoformat(self.session_data["end_time"]) - \
                   datetime.datetime.fromisoformat(self.session_data["start_time"])
        self.session_data["summary"]["session_duration_seconds"] = duration.total_seconds()
        
        # Count response types
        real_ai_responses = sum(1 for i in self.session_data["interactions"] 
                               if i["metadata"]["has_real_ai"])
        placeholder_responses = sum(1 for i in self.session_data["interactions"] 
                                   if i["metadata"]["is_placeholder"])
        
        self.session_data["summary"]["real_ai_responses"] = real_ai_responses
        self.session_data["summary"]["placeholder_responses"] = placeholder_responses
        
        # Save session summary
        with open(self.session_log, 'w', encoding='utf-8') as f:
            json.dump(self.session_data, f, indent=2, ensure_ascii=False)
        
        # Create human-readable summary
        summary_file = self.session_dir / "summary.md"
        self._create_markdown_summary(summary_file)
        
        print(f"✅ AI logging session finalized: {self.session_dir}")
        print(f"📊 Summary: {len(self.session_data['interactions'])} interactions, "
              f"{real_ai_responses} real AI responses, {placeholder_responses} placeholders")
        
        return self.session_dir
    
    def _create_markdown_summary(self, output_file: Path):
        """Create a human-readable markdown summary"""
        content = f"""# AI Processing Session Summary

**Session ID:** {self.session_data['session_id']}  
**Start Time:** {self.session_data['start_time']}  
**End Time:** {self.session_data.get('end_time', 'In Progress')}  
**Duration:** {self.session_data['summary'].get('session_duration_seconds', 0):.1f} seconds  

## Summary Statistics

- **Total Interactions:** {self.session_data['summary']['total_interactions']}
- **Successful Responses:** {self.session_data['summary']['successful_responses']}
- **Errors:** {self.session_data['summary']['errors']}
- **Real AI Responses:** {self.session_data['summary'].get('real_ai_responses', 0)}
- **Placeholder Responses:** {self.session_data['summary'].get('placeholder_responses', 0)}

## Agents Used

{', '.join(self.session_data['summary']['agents_used'])}

## Keys Processed

{', '.join(self.session_data['summary']['keys_processed'])}

## Detailed Interactions

"""
        
        for interaction in self.session_data['interactions']:
            content += f"""
### {interaction['agent_name']} - {interaction['key']}

**Time:** {interaction['timestamp']}  
**Processing Time:** {interaction['processing_time_seconds']:.2f}s  
**Success:** {'✅' if interaction['output']['success'] else '❌'}  
**Response Type:** {'🤖 Real AI' if interaction['metadata']['has_real_ai'] else '🔄 Placeholder'}  

**System Prompt Length:** {interaction['input']['system_prompt_length']} chars  
**User Prompt Length:** {interaction['input']['user_prompt_length']} chars  
**Response Length:** {interaction['output']['response_length']} chars  

**System Prompt:**
```
{interaction['input']['system_prompt'][:200]}...
```

**User Prompt:**
```
{interaction['input']['user_prompt'][:200]}...
```

**AI Response:**
```
{interaction['output']['ai_response'][:300]}...
```

---
"""
        
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(content)
    
    def get_session_info(self) -> Dict[str, Any]:
        """Get current session information"""
        return {
            "session_dir": str(self.session_dir),
            "interactions_count": len(self.session_data['interactions']),
            "session_id": self.session_data['session_id'],
            "start_time": self.session_data['start_time']
        }

# Global logger instance
_current_logger: Optional[AILogger] = None

def get_ai_logger() -> AILogger:
    """Get or create the current AI logger instance"""
    global _current_logger
    if _current_logger is None:
        _current_logger = AILogger()
    return _current_logger

def start_new_ai_session():
    """Start a new AI logging session"""
    global _current_logger
    if _current_logger:
        _current_logger.finalize_session()
    _current_logger = AILogger()
    return _current_logger

def finalize_current_session():
    """Finalize the current AI logging session"""
    global _current_logger
    if _current_logger:
        session_dir = _current_logger.finalize_session()
        _current_logger = None
        return session_dir
    return None

if __name__ == "__main__":
    # Test the logger
    logger = AILogger()
    
    # Test interaction
    logger.log_ai_interaction(
        agent_name="Agent 1",
        key="Cash",
        system_prompt="You are a financial analyst...",
        user_prompt="Analyze this cash data...",
        ai_response="Based on the analysis...",
        entity_name="Test Entity",
        processing_time=1.5
    )
    
    session_dir = logger.finalize_session()
    print(f"Test logging completed: {session_dir}") 
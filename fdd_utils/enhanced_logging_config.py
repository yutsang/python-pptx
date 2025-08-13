"""
Enhanced AI Agent Logging Configuration
Complements the existing AIAgentLogger with additional monitoring capabilities
"""

import logging
import json
import time
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional

class EnhancedAgentMonitor:
    """Enhanced monitoring for AI agent performance and interactions"""
    
    def __init__(self, base_logger=None):
        self.base_logger = base_logger
        self.session_start = time.time()
        self.agent_metrics = {
            'agent1': {'calls': 0, 'total_time': 0.0, 'avg_time': 0.0, 'errors': 0},
            'agent2': {'calls': 0, 'total_time': 0.0, 'avg_time': 0.0, 'errors': 0},
            'agent3': {'calls': 0, 'total_time': 0.0, 'avg_time': 0.0, 'errors': 0}
        }
        self.detailed_logs = []
        
        # Create enhanced logging directory
        self.log_dir = Path("logging/enhanced")
        self.log_dir.mkdir(parents=True, exist_ok=True)
        
        # Performance log file
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        self.performance_log = self.log_dir / f"agent_performance_{timestamp}.json"
        self.interaction_log = self.log_dir / f"detailed_interactions_{timestamp}.jsonl"
        
    def log_agent_start(self, agent_name: str, key: str, input_data: Dict):
        """Log when an agent starts processing"""
        log_entry = {
            'timestamp': datetime.now().isoformat(),
            'event': 'agent_start',
            'agent': agent_name,
            'key': key,
            'input_size': len(str(input_data)),
            'session_time': time.time() - self.session_start
        }
        self._write_interaction_log(log_entry)
        return time.time()  # Return start time for duration calculation
    
    def log_agent_complete(self, agent_name: str, key: str, start_time: float, 
                          output_data: Dict, success: bool = True):
        """Log when an agent completes processing"""
        duration = time.time() - start_time
        
        # Update metrics
        if agent_name in self.agent_metrics:
            metrics = self.agent_metrics[agent_name]
            metrics['calls'] += 1
            metrics['total_time'] += duration
            metrics['avg_time'] = metrics['total_time'] / metrics['calls']
            if not success:
                metrics['errors'] += 1
        
        log_entry = {
            'timestamp': datetime.now().isoformat(),
            'event': 'agent_complete',
            'agent': agent_name,
            'key': key,
            'duration': duration,
            'output_size': len(str(output_data)),
            'success': success,
            'session_time': time.time() - self.session_start
        }
        self._write_interaction_log(log_entry)
        self._update_performance_log()
    
    def log_agent_error(self, agent_name: str, key: str, error: str):
        """Log agent errors with context"""
        if agent_name in self.agent_metrics:
            self.agent_metrics[agent_name]['errors'] += 1
        
        log_entry = {
            'timestamp': datetime.now().isoformat(),
            'event': 'agent_error',
            'agent': agent_name,
            'key': key,
            'error': error,
            'session_time': time.time() - self.session_start
        }
        self._write_interaction_log(log_entry)
        self._update_performance_log()
    
    def _write_interaction_log(self, log_entry: Dict):
        """Write interaction log in JSONL format"""
        try:
            with open(self.interaction_log, 'a', encoding='utf-8') as f:
                f.write(json.dumps(log_entry) + '\n')
        except Exception as e:
            print(f"Failed to write interaction log: {e}")
    
    def _update_performance_log(self):
        """Update performance metrics file"""
        try:
            performance_data = {
                'session_start': datetime.fromtimestamp(self.session_start).isoformat(),
                'last_update': datetime.now().isoformat(),
                'session_duration': time.time() - self.session_start,
                'agent_metrics': self.agent_metrics,
                'total_agent_calls': sum(m['calls'] for m in self.agent_metrics.values()),
                'total_errors': sum(m['errors'] for m in self.agent_metrics.values())
            }
            
            with open(self.performance_log, 'w', encoding='utf-8') as f:
                json.dump(performance_data, f, indent=2)
        except Exception as e:
            print(f"Failed to update performance log: {e}")
    
    def get_performance_summary(self) -> Dict:
        """Get current performance summary"""
        total_calls = sum(m['calls'] for m in self.agent_metrics.values())
        total_errors = sum(m['errors'] for m in self.agent_metrics.values())
        
        return {
            'session_duration': time.time() - self.session_start,
            'total_agent_calls': total_calls,
            'total_errors': total_errors,
            'success_rate': (total_calls - total_errors) / total_calls if total_calls > 0 else 0,
            'agent_metrics': self.agent_metrics.copy()
        }
    
    def export_session_report(self) -> str:
        """Export comprehensive session report"""
        try:
            summary = self.get_performance_summary()
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            report_file = self.log_dir / f"session_report_{timestamp}.json"
            
            # Read all interaction logs
            interactions = []
            if self.interaction_log.exists():
                with open(self.interaction_log, 'r', encoding='utf-8') as f:
                    for line in f:
                        try:
                            interactions.append(json.loads(line.strip()))
                        except json.JSONDecodeError:
                            continue
            
            report_data = {
                'session_summary': summary,
                'detailed_interactions': interactions,
                'generated_at': datetime.now().isoformat()
            }
            
            with open(report_file, 'w', encoding='utf-8') as f:
                json.dump(report_data, f, indent=2)
            
            return str(report_file)
        except Exception as e:
            print(f"Failed to export session report: {e}")
            return ""

class AgentCommunicationLogger:
    """Log communication between agents and data flow"""
    
    def __init__(self):
        self.log_dir = Path("logging/communication")
        self.log_dir.mkdir(parents=True, exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        self.comm_log = self.log_dir / f"agent_communication_{timestamp}.jsonl"
    
    def log_data_flow(self, from_agent: str, to_agent: str, key: str, 
                     data_type: str, data_size: int):
        """Log data flow between agents"""
        log_entry = {
            'timestamp': datetime.now().isoformat(),
            'from_agent': from_agent,
            'to_agent': to_agent,
            'key': key,
            'data_type': data_type,
            'data_size': data_size,
            'flow_id': f"{from_agent}_{to_agent}_{key}"
        }
        
        try:
            with open(self.comm_log, 'a', encoding='utf-8') as f:
                f.write(json.dumps(log_entry) + '\n')
        except Exception as e:
            print(f"Failed to log communication: {e}")
    
    def log_content_modification(self, agent: str, key: str, original_size: int, 
                               modified_size: int, changes: List[str]):
        """Log when an agent modifies content"""
        log_entry = {
            'timestamp': datetime.now().isoformat(),
            'agent': agent,
            'key': key,
            'original_size': original_size,
            'modified_size': modified_size,
            'size_change': modified_size - original_size,
            'changes': changes,
            'change_count': len(changes)
        }
        
        try:
            with open(self.comm_log, 'a', encoding='utf-8') as f:
                f.write(json.dumps(log_entry) + '\n')
        except Exception as e:
            print(f"Failed to log content modification: {e}")

# Integration functions for existing system
def enhance_existing_logger(existing_logger):
    """Enhance the existing AIAgentLogger with additional monitoring"""
    if not hasattr(existing_logger, 'enhanced_monitor'):
        existing_logger.enhanced_monitor = EnhancedAgentMonitor(existing_logger)
        existing_logger.comm_logger = AgentCommunicationLogger()
    
    return existing_logger

def create_logging_dashboard_data():
    """Create data for a logging dashboard"""
    log_dir = Path("logging")
    if not log_dir.exists():
        return {}
    
    # Collect all log files
    log_files = list(log_dir.glob("**/*.json")) + list(log_dir.glob("**/*.jsonl"))
    
    dashboard_data = {
        'total_log_files': len(log_files),
        'log_directories': [str(d.relative_to(log_dir)) for d in log_dir.iterdir() if d.is_dir()],
        'recent_sessions': [],
        'agent_stats': {'agent1': 0, 'agent2': 0, 'agent3': 0}
    }
    
    # Parse recent session data
    for log_file in sorted(log_files, key=lambda x: x.stat().st_mtime, reverse=True)[:10]:
        try:
            if log_file.suffix == '.json':
                with open(log_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    dashboard_data['recent_sessions'].append({
                        'file': log_file.name,
                        'size': log_file.stat().st_size,
                        'modified': datetime.fromtimestamp(log_file.stat().st_mtime).isoformat()
                    })
        except Exception:
            continue
    
    return dashboard_data 
"""Processing Result domain model."""

from dataclasses import dataclass
from datetime import datetime
from typing import Dict, List, Optional, Any
from enum import Enum


class ProcessingStatus(Enum):
    """Status of processing operations."""
    PENDING = "pending"
    IN_PROGRESS = "in_progress"
    COMPLETED = "completed"
    FAILED = "failed"
    CANCELLED = "cancelled"


class AgentType(Enum):
    """Types of AI agents."""
    CONTENT_GENERATION = "content_generation"
    DATA_VALIDATION = "data_validation"
    PATTERN_COMPLIANCE = "pattern_compliance"


@dataclass
class AgentResult:
    """Result from a single AI agent."""
    agent_type: AgentType
    financial_key: str
    content: str
    processing_time: float
    status: ProcessingStatus
    metadata: Dict[str, Any]
    issues: List[str]
    corrections_made: bool
    timestamp: datetime
    
    def is_successful(self) -> bool:
        """Check if agent processing was successful."""
        return self.status == ProcessingStatus.COMPLETED and not self.issues


@dataclass 
class ProcessingResult:
    """Result of the complete AI processing pipeline."""
    
    entity_name: str
    statement_type: str
    agent_results: Dict[str, List[AgentResult]]  # key -> list of agent results
    final_content: Dict[str, str]  # key -> final content
    
    # Processing metadata
    total_processing_time: float
    started_at: datetime
    completed_at: Optional[datetime]
    status: ProcessingStatus
    
    # Summary information
    total_keys_processed: int
    successful_keys: List[str]
    failed_keys: List[str]
    warnings: List[str]
    errors: List[str]
    
    def __post_init__(self):
        """Calculate derived fields after initialization."""
        self._calculate_summary()
    
    def _calculate_summary(self) -> None:
        """Calculate summary statistics."""
        self.total_keys_processed = len(self.agent_results)
        self.successful_keys = []
        self.failed_keys = []
        
        for key, results in self.agent_results.items():
            # Check if all agents for this key were successful
            all_successful = all(result.is_successful() for result in results)
            
            if all_successful:
                self.successful_keys.append(key)
            else:
                self.failed_keys.append(key)
    
    def get_agent_result(self, key: str, agent_type: AgentType) -> Optional[AgentResult]:
        """Get specific agent result for a key."""
        key_results = self.agent_results.get(key, [])
        
        for result in key_results:
            if result.agent_type == agent_type:
                return result
        
        return None
    
    def get_processing_summary(self) -> Dict[str, Any]:
        """Get comprehensive processing summary."""
        agent_stats = {}
        total_issues = 0
        total_corrections = 0
        
        for key, results in self.agent_results.items():
            for result in results:
                agent_name = result.agent_type.value
                if agent_name not in agent_stats:
                    agent_stats[agent_name] = {
                        'processed': 0,
                        'successful': 0,
                        'failed': 0,
                        'avg_processing_time': 0.0,
                        'total_time': 0.0
                    }
                
                agent_stats[agent_name]['processed'] += 1
                agent_stats[agent_name]['total_time'] += result.processing_time
                
                if result.is_successful():
                    agent_stats[agent_name]['successful'] += 1
                else:
                    agent_stats[agent_name]['failed'] += 1
                
                total_issues += len(result.issues)
                if result.corrections_made:
                    total_corrections += 1
        
        # Calculate averages
        for stats in agent_stats.values():
            if stats['processed'] > 0:
                stats['avg_processing_time'] = stats['total_time'] / stats['processed']
        
        return {
            'total_keys': self.total_keys_processed,
            'successful_keys': len(self.successful_keys),
            'failed_keys': len(self.failed_keys),
            'success_rate': len(self.successful_keys) / self.total_keys_processed if self.total_keys_processed > 0 else 0,
            'total_processing_time': self.total_processing_time,
            'total_issues_found': total_issues,
            'total_corrections_made': total_corrections,
            'agent_statistics': agent_stats,
            'status': self.status.value,
            'warnings_count': len(self.warnings),
            'errors_count': len(self.errors)
        }
    
    def has_errors(self) -> bool:
        """Check if processing has any errors."""
        return len(self.errors) > 0 or len(self.failed_keys) > 0
    
    def has_warnings(self) -> bool:
        """Check if processing has any warnings."""
        return len(self.warnings) > 0
    
    def is_completed_successfully(self) -> bool:
        """Check if processing completed successfully."""
        return (
            self.status == ProcessingStatus.COMPLETED and
            not self.has_errors() and
            len(self.successful_keys) == self.total_keys_processed
        )
    
    def add_agent_result(self, result: AgentResult) -> None:
        """Add a new agent result."""
        key = result.financial_key
        
        if key not in self.agent_results:
            self.agent_results[key] = []
        
        self.agent_results[key].append(result)
        self._calculate_summary()
    
    def add_warning(self, warning: str) -> None:
        """Add a warning message."""
        self.warnings.append(warning)
    
    def add_error(self, error: str) -> None:
        """Add an error message."""
        self.errors.append(error)
        if self.status != ProcessingStatus.FAILED:
            self.status = ProcessingStatus.FAILED
    
    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary representation."""
        return {
            'entity_name': self.entity_name,
            'statement_type': self.statement_type,
            'agent_results': {
                key: [
                    {
                        'agent_type': result.agent_type.value,
                        'financial_key': result.financial_key,
                        'content': result.content,
                        'processing_time': result.processing_time,
                        'status': result.status.value,
                        'metadata': result.metadata,
                        'issues': result.issues,
                        'corrections_made': result.corrections_made,
                        'timestamp': result.timestamp.isoformat()
                    }
                    for result in results
                ]
                for key, results in self.agent_results.items()
            },
            'final_content': self.final_content,
            'total_processing_time': self.total_processing_time,
            'started_at': self.started_at.isoformat(),
            'completed_at': self.completed_at.isoformat() if self.completed_at else None,
            'status': self.status.value,
            'summary': self.get_processing_summary(),
            'warnings': self.warnings,
            'errors': self.errors
        } 
"""Report Template domain model."""

from dataclasses import dataclass
from datetime import datetime
from typing import Dict, List, Optional, Any
from enum import Enum


class TemplateType(Enum):
    """Types of report templates."""
    BALANCE_SHEET = "balance_sheet"
    INCOME_STATEMENT = "income_statement"
    DUE_DILIGENCE = "due_diligence"
    COMPREHENSIVE = "comprehensive"


@dataclass
class PatternOption:
    """Individual pattern option for content generation."""
    id: str
    name: str
    template: str
    description: str
    required_fields: List[str]
    optional_fields: List[str]
    weight: float = 1.0  # For pattern selection scoring
    
    def calculate_match_score(self, available_data: Dict[str, Any]) -> float:
        """Calculate how well this pattern matches available data."""
        required_score = 0.0
        optional_score = 0.0
        
        # Required fields must be present
        for field in self.required_fields:
            if field in available_data and available_data[field] is not None:
                required_score += 1.0
            else:
                return 0.0  # Missing required field, pattern not viable
        
        # Optional fields add bonus score
        for field in self.optional_fields:
            if field in available_data and available_data[field] is not None:
                optional_score += 0.5
        
        # Normalize scores
        max_required = len(self.required_fields)
        max_optional = len(self.optional_fields) * 0.5
        
        if max_required > 0:
            required_score = required_score / max_required
        else:
            required_score = 1.0
        
        if max_optional > 0:
            optional_score = optional_score / max_optional
        else:
            optional_score = 0.0
        
        # Weighted combination
        total_score = (required_score * 0.8 + optional_score * 0.2) * self.weight
        
        return total_score


@dataclass
class ReportTemplate:
    """Domain entity for report templates and patterns."""
    
    template_id: str
    name: str
    description: str
    template_type: TemplateType
    version: str
    
    # Pattern mappings for each financial key
    patterns: Dict[str, List[PatternOption]]  # key -> list of pattern options
    
    # Template metadata
    created_at: datetime
    updated_at: datetime
    created_by: str
    is_active: bool = True
    
    # Template configuration
    default_currency: str = "USD"
    reporting_standards: Optional[List[str]] = None
    
    def __post_init__(self):
        """Initialize default values and validate template."""
        if self.reporting_standards is None:
            self.reporting_standards = []
        self._validate_template()
    
    def _validate_template(self) -> None:
        """Validate template configuration."""
        if not self.template_id or not self.template_id.strip():
            raise ValueError("Template ID cannot be empty")
        
        if not self.patterns:
            raise ValueError("Template must have at least one pattern")
        
        # Validate patterns
        for key, pattern_list in self.patterns.items():
            if not pattern_list:
                raise ValueError(f"Key '{key}' must have at least one pattern")
            
            for pattern in pattern_list:
                if not isinstance(pattern, PatternOption):
                    raise ValueError(f"Invalid pattern for key '{key}'")
    
    def get_patterns_for_key(self, key: str) -> List[PatternOption]:
        """Get pattern options for a specific financial key."""
        return self.patterns.get(key, [])
    
    def select_best_pattern(self, key: str, available_data: Dict[str, Any]) -> Optional[PatternOption]:
        """Select the best pattern for a key based on available data."""
        patterns = self.get_patterns_for_key(key)
        
        if not patterns:
            return None
        
        # Score each pattern
        scored_patterns = []
        for pattern in patterns:
            score = pattern.calculate_match_score(available_data)
            if score > 0:  # Only consider viable patterns
                scored_patterns.append((pattern, score))
        
        if not scored_patterns:
            return None
        
        # Sort by score (descending) and return best match
        scored_patterns.sort(key=lambda x: x[1], reverse=True)
        return scored_patterns[0][0]
    
    def get_available_keys(self) -> List[str]:
        """Get all available financial keys in this template."""
        return list(self.patterns.keys())
    
    def get_required_fields_for_key(self, key: str) -> List[str]:
        """Get all required fields for a specific key across all patterns."""
        patterns = self.get_patterns_for_key(key)
        required_fields = set()
        
        for pattern in patterns:
            required_fields.update(pattern.required_fields)
        
        return list(required_fields)
    
    def get_template_summary(self) -> Dict[str, Any]:
        """Get summary information about this template."""
        total_patterns = sum(len(patterns) for patterns in self.patterns.values())
        
        key_stats = {}
        for key, patterns in self.patterns.items():
            key_stats[key] = {
                'pattern_count': len(patterns),
                'required_fields': self.get_required_fields_for_key(key)
            }
        
        return {
            'template_id': self.template_id,
            'name': self.name,
            'template_type': self.template_type.value,
            'version': self.version,
            'total_keys': len(self.patterns),
            'total_patterns': total_patterns,
            'is_active': self.is_active,
            'default_currency': self.default_currency,
            'reporting_standards': self.reporting_standards,
            'key_statistics': key_stats,
            'created_at': self.created_at.isoformat(),
            'updated_at': self.updated_at.isoformat()
        }
    
    def add_pattern(self, key: str, pattern: PatternOption) -> None:
        """Add a new pattern for a specific key."""
        if key not in self.patterns:
            self.patterns[key] = []
        
        self.patterns[key].append(pattern)
        self.updated_at = datetime.now()
    
    def remove_pattern(self, key: str, pattern_id: str) -> bool:
        """Remove a pattern for a specific key."""
        if key not in self.patterns:
            return False
        
        original_count = len(self.patterns[key])
        self.patterns[key] = [p for p in self.patterns[key] if p.id != pattern_id]
        
        if len(self.patterns[key]) < original_count:
            self.updated_at = datetime.now()
            return True
        
        return False
    
    def update_pattern(self, key: str, pattern_id: str, updated_pattern: PatternOption) -> bool:
        """Update an existing pattern."""
        if key not in self.patterns:
            return False
        
        for i, pattern in enumerate(self.patterns[key]):
            if pattern.id == pattern_id:
                self.patterns[key][i] = updated_pattern
                self.updated_at = datetime.now()
                return True
        
        return False
    
    def clone(self, new_id: str, new_name: str) -> 'ReportTemplate':
        """Create a copy of this template with a new ID and name."""
        return ReportTemplate(
            template_id=new_id,
            name=new_name,
            description=f"Clone of {self.description}",
            template_type=self.template_type,
            version=self.version,
            patterns=self.patterns.copy(),
            created_at=datetime.now(),
            updated_at=datetime.now(),
            created_by=self.created_by,
            is_active=True,
            default_currency=self.default_currency,
            reporting_standards=self.reporting_standards.copy() if self.reporting_standards else []
        )
    
    def to_dict(self) -> Dict[str, Any]:
        """Convert template to dictionary representation."""
        return {
            'template_id': self.template_id,
            'name': self.name,
            'description': self.description,
            'template_type': self.template_type.value,
            'version': self.version,
            'patterns': {
                key: [
                    {
                        'id': pattern.id,
                        'name': pattern.name,
                        'template': pattern.template,
                        'description': pattern.description,
                        'required_fields': pattern.required_fields,
                        'optional_fields': pattern.optional_fields,
                        'weight': pattern.weight
                    }
                    for pattern in patterns
                ]
                for key, patterns in self.patterns.items()
            },
            'created_at': self.created_at.isoformat(),
            'updated_at': self.updated_at.isoformat(),
            'created_by': self.created_by,
            'is_active': self.is_active,
            'default_currency': self.default_currency,
            'reporting_standards': self.reporting_standards
        }
    
    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'ReportTemplate':
        """Create template from dictionary representation."""
        patterns = {}
        
        for key, pattern_list in data['patterns'].items():
            patterns[key] = [
                PatternOption(
                    id=p['id'],
                    name=p['name'],
                    template=p['template'],
                    description=p['description'],
                    required_fields=p['required_fields'],
                    optional_fields=p['optional_fields'],
                    weight=p.get('weight', 1.0)
                )
                for p in pattern_list
            ]
        
        return cls(
            template_id=data['template_id'],
            name=data['name'],
            description=data['description'],
            template_type=TemplateType(data['template_type']),
            version=data['version'],
            patterns=patterns,
            created_at=datetime.fromisoformat(data['created_at']),
            updated_at=datetime.fromisoformat(data['updated_at']),
            created_by=data['created_by'],
            is_active=data.get('is_active', True),
            default_currency=data.get('default_currency', 'USD'),
            reporting_standards=data.get('reporting_standards', [])
        ) 
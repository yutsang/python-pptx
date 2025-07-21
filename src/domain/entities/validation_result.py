"""Validation Result domain model."""

from dataclasses import dataclass, field
from datetime import datetime
from typing import Dict, List, Optional, Any
from enum import Enum


class ValidationSeverity(Enum):
    """Severity levels for validation issues."""
    INFO = "info"
    WARNING = "warning"
    ERROR = "error"
    CRITICAL = "critical"


@dataclass
class ValidationIssue:
    """Individual validation issue."""
    code: str
    message: str
    severity: ValidationSeverity
    field: Optional[str] = None
    value: Optional[Any] = None
    suggestion: Optional[str] = None
    
    def is_blocking(self) -> bool:
        """Check if this issue blocks processing."""
        return self.severity in [ValidationSeverity.ERROR, ValidationSeverity.CRITICAL]


@dataclass
class ValidationResult:
    """Result of data validation operations."""
    
    is_valid: bool
    issues: List[ValidationIssue] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)
    errors: List[str] = field(default_factory=list)
    
    # Validation metadata
    validated_at: datetime = field(default_factory=datetime.now)
    validation_type: str = "general"
    validator_version: str = "1.0.0"
    
    # Performance metrics
    validation_time: float = 0.0
    fields_validated: int = 0
    
    def add_issue(self, issue: ValidationIssue) -> None:
        """Add a validation issue."""
        self.issues.append(issue)
        
        # Update validity based on issue severity
        if issue.is_blocking():
            self.is_valid = False
            self.errors.append(issue.message)
        else:
            self.warnings.append(issue.message)
    
    def add_error(self, code: str, message: str, field: Optional[str] = None, 
                  value: Optional[Any] = None, suggestion: Optional[str] = None) -> None:
        """Add an error issue."""
        issue = ValidationIssue(
            code=code,
            message=message,
            severity=ValidationSeverity.ERROR,
            field=field,
            value=value,
            suggestion=suggestion
        )
        self.add_issue(issue)
    
    def add_warning(self, code: str, message: str, field: Optional[str] = None,
                   value: Optional[Any] = None, suggestion: Optional[str] = None) -> None:
        """Add a warning issue."""
        issue = ValidationIssue(
            code=code,
            message=message,
            severity=ValidationSeverity.WARNING,
            field=field,
            value=value,
            suggestion=suggestion
        )
        self.add_issue(issue)
    
    def get_issues_by_severity(self, severity: ValidationSeverity) -> List[ValidationIssue]:
        """Get issues filtered by severity."""
        return [issue for issue in self.issues if issue.severity == severity]
    
    def get_blocking_issues(self) -> List[ValidationIssue]:
        """Get all blocking issues."""
        return [issue for issue in self.issues if issue.is_blocking()]
    
    def has_blocking_issues(self) -> bool:
        """Check if there are any blocking issues."""
        return len(self.get_blocking_issues()) > 0
    
    def get_summary(self) -> Dict[str, Any]:
        """Get validation summary."""
        severity_counts = {}
        for severity in ValidationSeverity:
            severity_counts[severity.value] = len(self.get_issues_by_severity(severity))
        
        return {
            'is_valid': self.is_valid,
            'total_issues': len(self.issues),
            'blocking_issues': len(self.get_blocking_issues()),
            'severity_breakdown': severity_counts,
            'validation_time': self.validation_time,
            'fields_validated': self.fields_validated,
            'validation_type': self.validation_type,
            'validated_at': self.validated_at.isoformat()
        }
    
    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary representation."""
        return {
            'is_valid': self.is_valid,
            'issues': [
                {
                    'code': issue.code,
                    'message': issue.message,
                    'severity': issue.severity.value,
                    'field': issue.field,
                    'value': str(issue.value) if issue.value is not None else None,
                    'suggestion': issue.suggestion
                }
                for issue in self.issues
            ],
            'warnings': self.warnings,
            'errors': self.errors,
            'validated_at': self.validated_at.isoformat(),
            'validation_type': self.validation_type,
            'validator_version': self.validator_version,
            'validation_time': self.validation_time,
            'fields_validated': self.fields_validated,
            'summary': self.get_summary()
        }
    
    @classmethod
    def create_valid(cls, validation_type: str = "general") -> 'ValidationResult':
        """Create a valid validation result."""
        return cls(
            is_valid=True,
            validation_type=validation_type
        )
    
    @classmethod
    def create_invalid(cls, error_message: str, validation_type: str = "general") -> 'ValidationResult':
        """Create an invalid validation result with error."""
        result = cls(
            is_valid=False,
            validation_type=validation_type
        )
        result.add_error("VALIDATION_FAILED", error_message)
        return result


@dataclass
class FinancialValidationResult(ValidationResult):
    """Specialized validation result for financial data."""
    
    # Financial-specific validation data
    expected_total: Optional[float] = None
    actual_total: Optional[float] = None
    variance: Optional[float] = None
    variance_percentage: Optional[float] = None
    
    # Balance validation
    assets_total: Optional[float] = None
    liabilities_total: Optional[float] = None
    equity_total: Optional[float] = None
    balance_sheet_balanced: Optional[bool] = None
    
    def __post_init__(self):
        """Calculate derived fields after initialization."""
        self._calculate_variance()
        self._check_balance_sheet_equation()
    
    def _calculate_variance(self) -> None:
        """Calculate variance between expected and actual totals."""
        if self.expected_total is not None and self.actual_total is not None:
            self.variance = abs(self.actual_total - self.expected_total)
            
            if self.expected_total != 0:
                self.variance_percentage = (self.variance / abs(self.expected_total)) * 100
            else:
                self.variance_percentage = 0.0
    
    def _check_balance_sheet_equation(self) -> None:
        """Check if balance sheet equation holds: Assets = Liabilities + Equity."""
        if all(x is not None for x in [self.assets_total, self.liabilities_total, self.equity_total]):
            left_side = self.assets_total or 0.0
            right_side = (self.liabilities_total or 0.0) + (self.equity_total or 0.0)
            
            # Allow for small rounding differences
            tolerance = 0.01
            self.balance_sheet_balanced = abs(left_side - right_side) <= tolerance
            
            if not self.balance_sheet_balanced:
                self.add_error(
                    "BALANCE_SHEET_UNBALANCED",
                    f"Balance sheet equation doesn't balance: Assets={left_side:.2f}, Liabilities+Equity={right_side:.2f}",
                    suggestion="Check individual account values for accuracy"
                )
    
    def get_financial_summary(self) -> Dict[str, Any]:
        """Get financial validation summary."""
        base_summary = self.get_summary()
        
        financial_summary = {
            'expected_total': self.expected_total,
            'actual_total': self.actual_total,
            'variance': self.variance,
            'variance_percentage': self.variance_percentage,
            'assets_total': self.assets_total,
            'liabilities_total': self.liabilities_total,
            'equity_total': self.equity_total,
            'balance_sheet_balanced': self.balance_sheet_balanced
        }
        
        return {**base_summary, **financial_summary} 
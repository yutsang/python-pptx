"""Financial Entity domain model."""

from dataclasses import dataclass
from datetime import datetime
from typing import Dict, List, Optional
from enum import Enum


class EntityType(Enum):
    """Types of financial entities."""
    REAL_ESTATE = "real_estate"
    CORPORATION = "corporation"
    PARTNERSHIP = "partnership"


class StatementType(Enum):
    """Types of financial statements."""
    BALANCE_SHEET = "balance_sheet"
    INCOME_STATEMENT = "income_statement" 
    CASH_FLOW = "cash_flow"
    ALL = "all"


@dataclass
class FinancialData:
    """Container for financial data by key."""
    key: str
    value: float
    currency: str = "USD"
    reporting_date: Optional[datetime] = None
    source_sheet: Optional[str] = None
    metadata: Optional[Dict] = None


@dataclass
class FinancialEntity:
    """Core domain entity representing a financial entity."""
    
    name: str
    entity_type: EntityType
    financial_data: Dict[str, FinancialData]
    statement_type: StatementType
    reporting_date: datetime
    entity_keywords: List[str]
    
    # Metadata
    created_at: datetime
    updated_at: datetime
    data_source: Optional[str] = None
    
    def __post_init__(self):
        """Validate entity data after initialization."""
        self._validate_entity_data()
    
    def _validate_entity_data(self) -> None:
        """Validate financial entity data."""
        if not self.name or not self.name.strip():
            raise ValueError("Entity name cannot be empty")
        
        if not self.financial_data:
            raise ValueError("Financial data cannot be empty")
        
        # Validate each financial data entry
        for key, data in self.financial_data.items():
            if not isinstance(data, FinancialData):
                raise ValueError(f"Invalid financial data for key {key}")
    
    def get_financial_value(self, key: str) -> Optional[float]:
        """Get financial value for a specific key."""
        financial_data = self.financial_data.get(key)
        return financial_data.value if financial_data else None
    
    def get_total_assets(self) -> float:
        """Calculate total assets from balance sheet keys."""
        asset_keys = ["Cash", "AR", "Prepayments", "OR", "Other CA", "IP", "Other NCA"]
        total = 0.0
        
        for key in asset_keys:
            value = self.get_financial_value(key)
            if value is not None:
                total += value
        
        return total
    
    def get_total_liabilities(self) -> float:
        """Calculate total liabilities from balance sheet keys."""
        liability_keys = ["AP", "Taxes payable", "OP"]
        total = 0.0
        
        for key in liability_keys:
            value = self.get_financial_value(key)
            if value is not None:
                total += value
        
        return total
    
    def get_entity_specific_keys(self) -> List[str]:
        """Get keys applicable to this entity type and location."""
        # Business rule: Ningbo and Nanjing don't typically have Reserve
        if self.name in ["Ningbo", "Nanjing"]:
            return [key for key in self.financial_data.keys() if key != "Reserve"]
        
        return list(self.financial_data.keys())
    
    def validate_financial_consistency(self) -> List[str]:
        """Validate financial data consistency."""
        issues = []
        
        # Basic validation rules
        cash_value = self.get_financial_value("Cash")
        if cash_value is not None and cash_value < 0:
            issues.append("Cash cannot be negative")
        
        # Cross-validation
        ap_value = self.get_financial_value("AP")
        if cash_value and ap_value and ap_value > cash_value * 5:
            issues.append("Accounts Payable significantly higher than Cash")
        
        return issues
    
    def to_dict(self) -> Dict:
        """Convert entity to dictionary representation."""
        return {
            'name': self.name,
            'entity_type': self.entity_type.value,
            'statement_type': self.statement_type.value,
            'reporting_date': self.reporting_date.isoformat(),
            'financial_data': {
                key: {
                    'value': data.value,
                    'currency': data.currency,
                    'reporting_date': data.reporting_date.isoformat() if data.reporting_date else None,
                    'source_sheet': data.source_sheet,
                    'metadata': data.metadata
                }
                for key, data in self.financial_data.items()
            },
            'entity_keywords': self.entity_keywords,
            'created_at': self.created_at.isoformat(),
            'updated_at': self.updated_at.isoformat(),
            'data_source': self.data_source
        }
    
    @classmethod
    def from_dict(cls, data: Dict) -> 'FinancialEntity':
        """Create entity from dictionary representation."""
        financial_data = {}
        
        for key, value_data in data['financial_data'].items():
            financial_data[key] = FinancialData(
                key=key,
                value=value_data['value'],
                currency=value_data.get('currency', 'USD'),
                reporting_date=datetime.fromisoformat(value_data['reporting_date']) if value_data.get('reporting_date') else None,
                source_sheet=value_data.get('source_sheet'),
                metadata=value_data.get('metadata')
            )
        
        return cls(
            name=data['name'],
            entity_type=EntityType(data['entity_type']),
            financial_data=financial_data,
            statement_type=StatementType(data['statement_type']),
            reporting_date=datetime.fromisoformat(data['reporting_date']),
            entity_keywords=data['entity_keywords'],
            created_at=datetime.fromisoformat(data['created_at']),
            updated_at=datetime.fromisoformat(data['updated_at']),
            data_source=data.get('data_source')
        ) 
"""Financial Data Repository interface."""

from abc import ABC, abstractmethod
from typing import List, Optional, Dict, Any
from datetime import datetime

from ..entities.financial_entity import FinancialEntity, StatementType
from ..entities.validation_result import ValidationResult


class FinancialDataRepository(ABC):
    """Abstract repository for financial data operations."""
    
    @abstractmethod
    async def save_entity(self, entity: FinancialEntity) -> str:
        """Save a financial entity and return its ID."""
        pass
    
    @abstractmethod
    async def get_entity_by_id(self, entity_id: str) -> Optional[FinancialEntity]:
        """Retrieve a financial entity by ID."""
        pass
    
    @abstractmethod
    async def get_entity_by_name(self, name: str) -> Optional[FinancialEntity]:
        """Retrieve a financial entity by name."""
        pass
    
    @abstractmethod
    async def get_entities_by_type(self, statement_type: StatementType) -> List[FinancialEntity]:
        """Get entities filtered by statement type."""
        pass
    
    @abstractmethod
    async def update_entity(self, entity_id: str, entity: FinancialEntity) -> bool:
        """Update an existing financial entity."""
        pass
    
    @abstractmethod
    async def delete_entity(self, entity_id: str) -> bool:
        """Delete a financial entity."""
        pass
    
    @abstractmethod
    async def get_entities_by_date_range(self, start_date: datetime, end_date: datetime) -> List[FinancialEntity]:
        """Get entities within a date range."""
        pass
    
    @abstractmethod
    async def search_entities(self, search_criteria: Dict[str, Any]) -> List[FinancialEntity]:
        """Search entities based on criteria."""
        pass
    
    @abstractmethod
    async def get_entity_validation_history(self, entity_id: str) -> List[ValidationResult]:
        """Get validation history for an entity."""
        pass
    
    @abstractmethod
    async def save_validation_result(self, entity_id: str, validation_result: ValidationResult) -> str:
        """Save validation result for an entity."""
        pass 
"""Report Repository interface."""

from abc import ABC, abstractmethod
from typing import List, Optional, Dict, Any
from datetime import datetime

from ..entities.processing_result import ProcessingResult, ProcessingStatus


class ReportRepository(ABC):
    """Abstract repository for report operations."""
    
    @abstractmethod
    async def save_processing_result(self, result: ProcessingResult) -> str:
        """Save a processing result and return its ID."""
        pass
    
    @abstractmethod
    async def get_processing_result_by_id(self, result_id: str) -> Optional[ProcessingResult]:
        """Retrieve a processing result by ID."""
        pass
    
    @abstractmethod
    async def get_processing_results_by_entity(self, entity_name: str) -> List[ProcessingResult]:
        """Get all processing results for an entity."""
        pass
    
    @abstractmethod
    async def get_processing_results_by_status(self, status: ProcessingStatus) -> List[ProcessingResult]:
        """Get processing results filtered by status."""
        pass
    
    @abstractmethod
    async def update_processing_status(self, result_id: str, status: ProcessingStatus) -> bool:
        """Update the status of a processing result."""
        pass
    
    @abstractmethod
    async def get_processing_results_by_date_range(self, start_date: datetime, end_date: datetime) -> List[ProcessingResult]:
        """Get processing results within a date range."""
        pass
    
    @abstractmethod
    async def delete_processing_result(self, result_id: str) -> bool:
        """Delete a processing result."""
        pass
    
    @abstractmethod
    async def get_processing_statistics(self, entity_name: Optional[str] = None) -> Dict[str, Any]:
        """Get processing statistics, optionally filtered by entity."""
        pass
    
    @abstractmethod
    async def get_failed_processing_results(self, days_back: int = 7) -> List[ProcessingResult]:
        """Get recent failed processing results for analysis."""
        pass
    
    @abstractmethod
    async def cleanup_old_results(self, days_to_keep: int = 30) -> int:
        """Clean up old processing results and return count of deleted items."""
        pass 
"""Template Repository interface."""

from abc import ABC, abstractmethod
from typing import List, Optional, Dict, Any

from ..entities.report_template import ReportTemplate, TemplateType


class TemplateRepository(ABC):
    """Abstract repository for report template operations."""
    
    @abstractmethod
    async def save_template(self, template: ReportTemplate) -> str:
        """Save a report template and return its ID."""
        pass
    
    @abstractmethod
    async def get_template_by_id(self, template_id: str) -> Optional[ReportTemplate]:
        """Retrieve a template by ID."""
        pass
    
    @abstractmethod
    async def get_templates_by_type(self, template_type: TemplateType) -> List[ReportTemplate]:
        """Get templates filtered by type."""
        pass
    
    @abstractmethod
    async def get_active_templates(self) -> List[ReportTemplate]:
        """Get all active templates."""
        pass
    
    @abstractmethod
    async def update_template(self, template_id: str, template: ReportTemplate) -> bool:
        """Update an existing template."""
        pass
    
    @abstractmethod
    async def deactivate_template(self, template_id: str) -> bool:
        """Deactivate a template (soft delete)."""
        pass
    
    @abstractmethod
    async def delete_template(self, template_id: str) -> bool:
        """Permanently delete a template."""
        pass
    
    @abstractmethod
    async def get_template_versions(self, template_id: str) -> List[ReportTemplate]:
        """Get all versions of a template."""
        pass
    
    @abstractmethod
    async def get_default_template(self, template_type: TemplateType) -> Optional[ReportTemplate]:
        """Get the default template for a given type."""
        pass
    
    @abstractmethod
    async def search_templates(self, search_criteria: Dict[str, Any]) -> List[ReportTemplate]:
        """Search templates based on criteria."""
        pass 
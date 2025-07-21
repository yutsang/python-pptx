"""Request DTOs for application layer."""

from dataclasses import dataclass
from typing import Optional, Dict, Any, BinaryIO
from datetime import datetime

from ...domain.entities.financial_entity import StatementType, EntityType


@dataclass
class ProcessFinancialDataRequest:
    """Request to process financial data from Excel file."""
    
    entity_name: str
    statement_type: StatementType
    excel_file_data: bytes
    excel_filename: str
    
    # Optional configuration
    entity_type: EntityType = EntityType.REAL_ESTATE
    entity_keywords: Optional[list[str]] = None
    template_id: Optional[str] = None
    ai_model: str = "gpt-4o-mini"
    
    # Processing options
    validate_data: bool = True
    check_patterns: bool = True
    generate_report: bool = True
    
    # Metadata
    uploaded_by: Optional[str] = None
    upload_timestamp: Optional[datetime] = None
    
    def __post_init__(self):
        """Set default values and validate request."""
        if self.upload_timestamp is None:
            self.upload_timestamp = datetime.now()
        
        if self.entity_keywords is None:
            self.entity_keywords = [self.entity_name]
        
        self._validate_request()
    
    def _validate_request(self) -> None:
        """Validate request data."""
        if not self.entity_name or not self.entity_name.strip():
            raise ValueError("Entity name is required")
        
        if not self.excel_file_data:
            raise ValueError("Excel file data is required")
        
        if not self.excel_filename:
            raise ValueError("Excel filename is required")


@dataclass
class GenerateReportRequest:
    """Request to generate PowerPoint report."""
    
    entity_name: str
    statement_type: StatementType
    processing_result_id: str
    template_id: Optional[str] = None
    
    # Report configuration
    project_name: Optional[str] = None
    include_summary: bool = True
    include_details: bool = True
    
    # Output options
    output_format: str = "pptx"
    output_filename: Optional[str] = None
    
    # Metadata
    requested_by: Optional[str] = None
    request_timestamp: Optional[datetime] = None
    
    def __post_init__(self):
        """Set default values and validate request."""
        if self.request_timestamp is None:
            self.request_timestamp = datetime.now()
        
        if self.project_name is None:
            self.project_name = self.entity_name
        
        if self.output_filename is None:
            timestamp = self.request_timestamp.strftime("%Y%m%d_%H%M%S")
            self.output_filename = f"{self.project_name}_{self.statement_type.value}_{timestamp}.pptx"
        
        self._validate_request()
    
    def _validate_request(self) -> None:
        """Validate request data."""
        if not self.entity_name:
            raise ValueError("Entity name is required")
        
        if not self.processing_result_id:
            raise ValueError("Processing result ID is required")


@dataclass
class ValidateFinancialDataRequest:
    """Request to validate financial data."""
    
    entity_name: str
    financial_data: Dict[str, Any]
    statement_type: StatementType
    
    # Validation options
    strict_validation: bool = False
    check_balance_equation: bool = True
    validate_cross_references: bool = True
    
    # Expected totals for validation
    expected_totals: Optional[Dict[str, float]] = None
    
    # Metadata
    validation_requested_by: Optional[str] = None
    validation_timestamp: Optional[datetime] = None
    
    def __post_init__(self):
        """Set default values and validate request."""
        if self.validation_timestamp is None:
            self.validation_timestamp = datetime.now()
        
        self._validate_request()
    
    def _validate_request(self) -> None:
        """Validate request data."""
        if not self.entity_name:
            raise ValueError("Entity name is required")
        
        if not self.financial_data:
            raise ValueError("Financial data is required")


@dataclass
class ProcessingStatusRequest:
    """Request to get processing status."""
    
    processing_id: str
    include_details: bool = False
    
    def _validate_request(self) -> None:
        """Validate request data."""
        if not self.processing_id:
            raise ValueError("Processing ID is required") 
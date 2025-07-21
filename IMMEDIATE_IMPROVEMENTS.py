# IMMEDIATE IMPROVEMENTS - Quick Wins for Current System

# 1. EXTRACT BUSINESS LOGIC FROM UI
# ============================================

# Current Issue: Business logic mixed with Streamlit UI code
# Quick Fix: Create service layer

class FinancialDataService:
    """Extract financial data processing logic"""
    
    def __init__(self, config_manager, cache_manager):
        self.config = config_manager
        self.cache = cache_manager
    
    def process_entity_data(self, uploaded_file, entity_name: str, entity_helpers: str, statement_type: str):
        """Centralized business logic for processing financial data"""
        try:
            # Validate inputs
            self._validate_inputs(uploaded_file, entity_name, statement_type)
            
            # Process entity configuration
            entity_suffixes = [s.strip() for s in entity_helpers.split(',') if s.strip()]
            entity_keywords = [f"{entity_name} {suffix}" for suffix in entity_suffixes if suffix]
            if not entity_keywords:
                entity_keywords = [entity_name]
            
            # Load configurations
            config, mapping, pattern, prompts = self._load_configurations()
            
            # Extract worksheet sections
            sections_by_key = self._extract_worksheet_sections(
                uploaded_file, mapping, entity_name, entity_suffixes
            )
            
            # Filter keys by statement type
            filtered_keys = self._filter_keys_by_statement_type(sections_by_key, statement_type)
            
            return ProcessingResult(
                entity_name=entity_name,
                entity_keywords=entity_keywords,
                sections_by_key=sections_by_key,
                filtered_keys=filtered_keys,
                config=config,
                mapping=mapping,
                pattern=pattern
            )
            
        except Exception as e:
            raise FinancialDataProcessingError(f"Failed to process entity data: {e}") from e
    
    def _validate_inputs(self, uploaded_file, entity_name: str, statement_type: str):
        """Validate input parameters"""
        if not uploaded_file:
            raise ValueError("Uploaded file is required")
        
        valid_entities = ["Haining", "Nanjing", "Ningbo"]
        if entity_name not in valid_entities:
            raise ValueError(f"Entity must be one of: {valid_entities}")
        
        valid_types = ["BS", "IS", "ALL"]
        if statement_type not in valid_types:
            raise ValueError(f"Statement type must be one of: {valid_types}")

# 2. CONFIGURATION MANAGEMENT
# ============================================

class ConfigurationManager:
    """Centralized configuration management"""
    
    def __init__(self, config_dir: str = "utils"):
        self.config_dir = Path(config_dir)
        self._cache = {}
    
    def get_ai_config(self) -> dict:
        """Get AI provider configuration with validation"""
        if 'ai_config' not in self._cache:
            config_path = self.config_dir / "config.json"
            
            if not config_path.exists():
                raise ConfigurationError(f"AI config file not found: {config_path}")
            
            with open(config_path, 'r') as f:
                config = json.load(f)
            
            # Validate required fields
            required_fields = ['OPENAI_API_KEY', 'OPENAI_API_BASE', 'CHAT_MODEL']
            missing_fields = [field for field in required_fields if not config.get(field)]
            
            if missing_fields:
                raise ConfigurationError(f"Missing required AI config fields: {missing_fields}")
            
            self._cache['ai_config'] = config
        
        return self._cache['ai_config']
    
    def get_entity_config(self) -> dict:
        """Get entity-specific configuration"""
        return {
            "supported_entities": ["Haining", "Nanjing", "Ningbo"],
            "balance_sheet_keys": [
                "Cash", "AR", "Prepayments", "OR", "Other CA", "IP", "Other NCA",
                "AP", "Taxes payable", "OP", "Capital", "Reserve"
            ],
            "income_statement_keys": [
                "OI", "OC", "Tax and Surcharges", "GA", "Fin Exp", "Cr Loss", 
                "Other Income", "Non-operating Income", "Non-operating Exp", 
                "Income tax", "LT DTA"
            ],
            "entity_specific_rules": {
                "Ningbo": {"exclude_keys": ["Reserve"]},
                "Nanjing": {"exclude_keys": ["Reserve"]}
            }
        }

# 3. IMPROVED ERROR HANDLING
# ============================================

class DueDiligenceError(Exception):
    """Base exception for due diligence processing"""
    pass

class ConfigurationError(DueDiligenceError):
    """Configuration-related errors"""
    pass

class FinancialDataProcessingError(DueDiligenceError):
    """Financial data processing errors"""
    pass

class AIProcessingError(DueDiligenceError):
    """AI processing errors"""
    pass

class ErrorHandler:
    """Centralized error handling with user-friendly messages"""
    
    @staticmethod
    def handle_processing_error(error: Exception, context: str = ""):
        """Convert technical errors to user-friendly messages"""
        error_mapping = {
            ConfigurationError: "Configuration issue. Please check your settings.",
            FinancialDataProcessingError: "Failed to process financial data. Please check your Excel file format.",
            AIProcessingError: "AI service temporarily unavailable. Please try again later.",
            FileNotFoundError: "Required file not found. Please check file paths.",
            PermissionError: "Permission denied. Please check file permissions.",
            ValueError: "Invalid input data. Please verify your inputs."
        }
        
        error_type = type(error)
        user_message = error_mapping.get(error_type, "An unexpected error occurred.")
        
        # Log technical details
        logging.error(f"Error in {context}: {error_type.__name__}: {str(error)}")
        
        return f"{user_message} (Context: {context})" if context else user_message

# 4. AI AGENT FACTORY PATTERN
# ============================================

class AIAgentFactory:
    """Factory for creating AI agents with proper configuration"""
    
    def __init__(self, config_manager: ConfigurationManager):
        self.config_manager = config_manager
    
    def create_content_agent(self) -> 'ContentGenerationAgent':
        """Create content generation agent"""
        ai_config = self.config_manager.get_ai_config()
        
        return ContentGenerationAgent(
            api_key=ai_config['OPENAI_API_KEY'],
            api_base=ai_config['OPENAI_API_BASE'],
            model=ai_config['CHAT_MODEL'],
            max_retries=3,
            timeout=30
        )
    
    def create_validation_agent(self) -> 'DataValidationAgent':
        """Create data validation agent"""
        ai_config = self.config_manager.get_ai_config()
        
        return DataValidationAgent(
            api_key=ai_config['OPENAI_API_KEY'],
            api_base=ai_config['OPENAI_API_BASE'],
            model=ai_config['CHAT_MODEL'],
            validation_rules=self._load_validation_rules()
        )
    
    def create_pattern_agent(self) -> 'PatternValidationAgent':
        """Create pattern validation agent"""
        ai_config = self.config_manager.get_ai_config()
        
        return PatternValidationAgent(
            api_key=ai_config['OPENAI_API_KEY'],
            api_base=ai_config['OPENAI_API_BASE'],
            model=ai_config['CHAT_MODEL'],
            patterns=self._load_patterns()
        )

# 5. IMPROVED CACHING STRATEGY
# ============================================

class IntelligentCacheManager:
    """Improved caching with TTL and intelligent invalidation"""
    
    def __init__(self, cache_dir: str = "cache"):
        self.cache_dir = Path(cache_dir)
        self.cache_dir.mkdir(exist_ok=True)
        self.memory_cache = {}
        self.ttl_cache = {}
    
    def get_processed_excel(self, file_hash: str, entity_name: str, ttl_hours: int = 24):
        """Get cached Excel processing result with TTL"""
        cache_key = f"excel_{file_hash}_{entity_name}"
        
        # Check memory cache first
        if cache_key in self.memory_cache:
            cache_time = self.ttl_cache.get(cache_key, 0)
            if time.time() - cache_time < ttl_hours * 3600:
                return self.memory_cache[cache_key]
            else:
                # Expired, remove from cache
                del self.memory_cache[cache_key]
                del self.ttl_cache[cache_key]
        
        # Check disk cache
        cache_file = self.cache_dir / f"{cache_key}.json"
        if cache_file.exists():
            try:
                file_age = time.time() - cache_file.stat().st_mtime
                if file_age < ttl_hours * 3600:
                    with open(cache_file, 'r') as f:
                        data = json.load(f)
                    
                    # Load into memory cache
                    self.memory_cache[cache_key] = data
                    self.ttl_cache[cache_key] = time.time()
                    return data
            except Exception:
                # Invalid cache file, remove it
                cache_file.unlink(missing_ok=True)
        
        return None
    
    def set_processed_excel(self, file_hash: str, entity_name: str, data: dict):
        """Cache Excel processing result"""
        cache_key = f"excel_{file_hash}_{entity_name}"
        
        # Save to memory cache
        self.memory_cache[cache_key] = data
        self.ttl_cache[cache_key] = time.time()
        
        # Save to disk cache
        cache_file = self.cache_dir / f"{cache_key}.json"
        with open(cache_file, 'w') as f:
            json.dump(data, f, default=str)

# 6. VALIDATION FRAMEWORK
# ============================================

class DataValidator:
    """Data validation framework for financial data"""
    
    def __init__(self):
        self.validation_rules = self._load_validation_rules()
    
    def validate_financial_data(self, entity_data: dict) -> ValidationResult:
        """Validate financial data against business rules"""
        errors = []
        warnings = []
        
        # Basic data structure validation
        required_keys = ["entity_name", "financial_data", "statement_type"]
        for key in required_keys:
            if key not in entity_data:
                errors.append(f"Missing required field: {key}")
        
        if errors:
            return ValidationResult(is_valid=False, errors=errors, warnings=warnings)
        
        # Financial data validation
        financial_data = entity_data["financial_data"]
        entity_name = entity_data["entity_name"]
        
        # Entity-specific validation
        if entity_name in ["Ningbo", "Nanjing"] and "Reserve" in financial_data:
            warnings.append(f"Reserve data found for {entity_name} but typically not applicable")
        
        # Balance sheet validation
        if "Cash" in financial_data:
            cash_value = financial_data["Cash"]
            if isinstance(cash_value, (int, float)) and cash_value < 0:
                errors.append("Cash cannot be negative")
        
        # Cross-validation rules
        if "AP" in financial_data and "Cash" in financial_data:
            try:
                ap_value = float(financial_data["AP"])
                cash_value = float(financial_data["Cash"])
                if ap_value > cash_value * 5:  # Arbitrary business rule
                    warnings.append("Accounts Payable significantly higher than Cash - please verify")
            except (ValueError, TypeError):
                warnings.append("Could not validate AP/Cash ratio - non-numeric values")
        
        return ValidationResult(
            is_valid=len(errors) == 0,
            errors=errors,
            warnings=warnings
        )
    
    def _load_validation_rules(self) -> dict:
        """Load validation rules from configuration"""
        # This would typically load from a configuration file
        return {
            "balance_sheet": {
                "required_keys": ["Cash", "AR"],
                "optional_keys": ["Prepayments", "OR", "Other CA"],
                "validation_rules": {
                    "Cash": {"min_value": 0, "type": "number"},
                    "AR": {"min_value": 0, "type": "number"}
                }
            }
        }

# 7. USAGE IN STREAMLIT (REFACTORED)
# ============================================

def refactored_main():
    """Refactored main function with better separation of concerns"""
    
    # Initialize services
    config_manager = ConfigurationManager()
    cache_manager = IntelligentCacheManager()
    financial_service = FinancialDataService(config_manager, cache_manager)
    ai_factory = AIAgentFactory(config_manager)
    validator = DataValidator()
    error_handler = ErrorHandler()
    
    st.title("📊 Financial Data Processor (Refactored)")
    
    try:
        # UI inputs (keep UI logic minimal)
        uploaded_file = st.file_uploader("Upload Excel File", type=['xlsx', 'xls'])
        entity_name = st.selectbox("Select Entity", ["Haining", "Nanjing", "Ningbo"])
        statement_type = st.radio("Statement Type", ["Balance Sheet", "Income Statement", "All"])
        
        if uploaded_file and st.button("Process with AI"):
            # Business logic handled by service layer
            with st.spinner("Processing financial data..."):
                try:
                    # Process data
                    result = financial_service.process_entity_data(
                        uploaded_file, entity_name, "Wanpu,Limited,", statement_type
                    )
                    
                    # Validate data
                    validation_result = validator.validate_financial_data({
                        "entity_name": entity_name,
                        "financial_data": result.sections_by_key,
                        "statement_type": statement_type
                    })
                    
                    if not validation_result.is_valid:
                        for error in validation_result.errors:
                            st.error(error)
                        return
                    
                    for warning in validation_result.warnings:
                        st.warning(warning)
                    
                    # Create AI agents
                    content_agent = ai_factory.create_content_agent()
                    validation_agent = ai_factory.create_validation_agent()
                    pattern_agent = ai_factory.create_pattern_agent()
                    
                    # Process with AI pipeline
                    ai_results = process_with_ai_pipeline(
                        [content_agent, validation_agent, pattern_agent],
                        result
                    )
                    
                    st.success("✅ Processing completed successfully!")
                    st.session_state['processing_result'] = result
                    st.session_state['ai_results'] = ai_results
                    
                except Exception as e:
                    error_message = error_handler.handle_processing_error(e, "AI Processing")
                    st.error(error_message)
                    
    except Exception as e:
        error_message = error_handler.handle_processing_error(e, "Application")
        st.error(error_message)

# 8. MONITORING AND METRICS
# ============================================

class MetricsCollector:
    """Simple metrics collection for current system"""
    
    def __init__(self):
        self.metrics = {
            'processing_times': [],
            'error_counts': {},
            'success_counts': 0,
            'entity_processing_counts': {}
        }
    
    def record_processing_time(self, entity_name: str, processing_time: float):
        """Record processing time metrics"""
        self.metrics['processing_times'].append({
            'entity': entity_name,
            'time': processing_time,
            'timestamp': datetime.now()
        })
        
        self.metrics['entity_processing_counts'][entity_name] = \
            self.metrics['entity_processing_counts'].get(entity_name, 0) + 1
        
        self.metrics['success_counts'] += 1
    
    def record_error(self, error_type: str, context: str):
        """Record error metrics"""
        error_key = f"{error_type}_{context}"
        self.metrics['error_counts'][error_key] = \
            self.metrics['error_counts'].get(error_key, 0) + 1
    
    def get_performance_summary(self) -> dict:
        """Get performance summary"""
        if not self.metrics['processing_times']:
            return {"message": "No processing data available"}
        
        times = [m['time'] for m in self.metrics['processing_times']]
        
        return {
            'avg_processing_time': sum(times) / len(times),
            'min_processing_time': min(times),
            'max_processing_time': max(times),
            'total_successful_runs': self.metrics['success_counts'],
            'total_errors': sum(self.metrics['error_counts'].values()),
            'entity_breakdown': self.metrics['entity_processing_counts']
        }

# Apply these improvements gradually:
# 1. Start with error handling and configuration management
# 2. Extract business logic from UI components
# 3. Implement improved caching
# 4. Add validation framework
# 5. Introduce metrics collection
# 6. Gradually refactor toward full architecture 
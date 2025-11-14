"""
AI Pipeline Orchestrator
Manages the execution of all AI agents in sequence or individually
Integrates with process_databook and content_generation modules
"""

import yaml
import os
from typing import Dict, List, Optional, Any
from fdd_utils.process_databook import extract_data_from_excel
from fdd_utils.content_generation import (
    ai_pipeline_full,
    ai_pipeline_agent_only
)


class AIPipelineOrchestrator:
    """
    Orchestrates the complete AI pipeline from data extraction to final output.
    """
    
    def __init__(self, config_path: str = 'fdd_utils/config.yml'):
        """
        Initialize the orchestrator.
        
        Args:
            config_path: Path to configuration file
        """
        self.config_path = config_path
        self.config = self._load_config()
        
        # Extract settings from config
        self.default_model = self.config.get('default', {}).get('ai_provider', 'deepseek')
        self.use_heuristic = self.config.get('default', {}).get('use_heuristic', False)
        self.processing_config = self.config.get('processing', {})
        self.use_multithreading = self.processing_config.get('use_multithreading', True)
        self.max_workers = self.processing_config.get('max_workers', 4)
        
        # Results storage
        self.dfs = None
        self.mapping_keys = None
        self.report_language = None
        self.results = {}
    
    def _load_config(self) -> Dict:
        """Load configuration from YAML file."""
        try:
            with open(self.config_path, 'r', encoding='utf-8') as f:
                return yaml.safe_load(f) or {}
        except FileNotFoundError:
            raise FileNotFoundError(f"Config file not found: {self.config_path}")
    
    def load_data(
        self, 
        databook_path: str, 
        entity_name: str, 
        mode: str = "All"
    ) -> Dict[str, Any]:
        """
        Load and extract data from Excel file.
        
        Args:
            databook_path: Path to Excel databook
            entity_name: Name of entity to extract
            mode: Filter mode for worksheets
        
        Returns:
            Dictionary with extraction results
        """
        print(f"Loading data from: {databook_path}")
        print(f"Entity: {entity_name}, Mode: {mode}")
        
        dfs, workbook_list, result_type, report_language = extract_data_from_excel(
            databook_path, entity_name, mode
        )
        
        self.dfs = dfs
        self.mapping_keys = workbook_list
        self.report_language = report_language or 'Eng'
        
        print(f"Extracted {len(workbook_list)} worksheets")
        print(f"Result type: {result_type}")
        print(f"Detected language: {self.report_language}")
        
        return {
            'dfs': dfs,
            'mapping_keys': workbook_list,
            'result_type': result_type,
            'report_language': self.report_language
        }
    
    def run_full_pipeline(
        self,
        model_type: Optional[str] = None,
        language: Optional[str] = None,
        use_heuristic: Optional[bool] = None,
        use_multithreading: Optional[bool] = None,
        max_workers: Optional[int] = None
    ) -> Dict[str, str]:
        """
        Run the full AI pipeline (all 4 agents in sequence).
        
        Args:
            model_type: AI model type (overrides config)
            language: Language for prompts (overrides detected)
            use_heuristic: Use heuristic mode (overrides config)
            use_multithreading: Use multi-threading (overrides config)
            max_workers: Max parallel workers (overrides config)
        
        Returns:
            Dictionary of final results by mapping key
        """
        if self.dfs is None or self.mapping_keys is None:
            raise ValueError("Data not loaded. Call load_data() first.")
        
        # Use provided values or fall back to config/detected
        model = model_type or self.default_model
        lang = language or self.report_language
        heuristic = use_heuristic if use_heuristic is not None else self.use_heuristic
        multithread = use_multithreading if use_multithreading is not None else self.use_multithreading
        workers = max_workers or self.max_workers
        
        print("\n" + "="*60)
        print("RUNNING FULL AI PIPELINE (4 Agents)")
        print("="*60)
        print(f"Model: {model}")
        print(f"Language: {lang}")
        print(f"Heuristic Mode: {heuristic}")
        print(f"Multi-threading: {multithread} (workers: {workers})")
        print(f"Items to process: {len(self.mapping_keys)}")
        print("="*60 + "\n")
        
        results = ai_pipeline_full(
            mapping_keys=self.mapping_keys,
            dfs=self.dfs,
            model_type=model,
            language=lang,
            use_heuristic=heuristic,
            use_multithreading=multithread,
            max_workers=workers
        )
        
        self.results = results
        
        print("\n" + "="*60)
        print("PIPELINE COMPLETED")
        print("="*60)
        print(f"Processed: {len(results)} items")
        print(f"Successful: {sum(1 for v in results.values() if v is not None)}")
        print(f"Failed: {sum(1 for v in results.values() if v is None)}")
        print("="*60 + "\n")
        
        return results
    
    def run_agent_only(
        self,
        agent_name: str,
        previous_results: Optional[Dict[str, str]] = None,
        model_type: Optional[str] = None,
        language: Optional[str] = None,
        use_heuristic: Optional[bool] = None,
        use_multithreading: Optional[bool] = None,
        max_workers: Optional[int] = None
    ) -> Dict[str, str]:
        """
        Run only a specific agent.
        
        Args:
            agent_name: Name of agent ('agent_1', 'agent_2', 'agent_3', 'agent_4')
            previous_results: Previous results (required for agents 2-4)
            model_type: AI model type (overrides config)
            language: Language for prompts (overrides detected)
            use_heuristic: Use heuristic mode (overrides config)
            use_multithreading: Use multi-threading (overrides config)
            max_workers: Max parallel workers (overrides config)
        
        Returns:
            Dictionary of results by mapping key
        """
        if self.dfs is None or self.mapping_keys is None:
            raise ValueError("Data not loaded. Call load_data() first.")
        
        # Validate agent name
        valid_agents = ['agent_1', 'agent_2', 'agent_3', 'agent_4']
        if agent_name not in valid_agents:
            raise ValueError(f"Invalid agent name: {agent_name}. Must be one of {valid_agents}")
        
        # Check for previous results if needed
        if agent_name != 'agent_1' and previous_results is None:
            raise ValueError(f"{agent_name} requires previous_results parameter")
        
        # Use provided values or fall back to config/detected
        model = model_type or self.default_model
        lang = language or self.report_language
        heuristic = use_heuristic if use_heuristic is not None else self.use_heuristic
        multithread = use_multithreading if use_multithreading is not None else self.use_multithreading
        workers = max_workers or self.max_workers
        
        print("\n" + "="*60)
        print(f"RUNNING {agent_name.upper()} ONLY")
        print("="*60)
        print(f"Model: {model}")
        print(f"Language: {lang}")
        print(f"Items to process: {len(self.mapping_keys)}")
        print("="*60 + "\n")
        
        results = ai_pipeline_agent_only(
            agent_name=agent_name,
            mapping_keys=self.mapping_keys,
            dfs=self.dfs,
            model_type=model,
            language=lang,
            use_heuristic=heuristic,
            use_multithreading=multithread,
            max_workers=workers,
            previous_results=previous_results
        )
        
        self.results = results
        
        print("\n" + "="*60)
        print(f"{agent_name.upper()} COMPLETED")
        print("="*60)
        print(f"Processed: {len(results)} items")
        print("="*60 + "\n")
        
        return results
    
    def run_sequential_agents(
        self,
        agents: List[str],
        model_type: Optional[str] = None,
        language: Optional[str] = None,
        use_heuristic: Optional[bool] = None,
        use_multithreading: Optional[bool] = None,
        max_workers: Optional[int] = None
    ) -> Dict[str, str]:
        """
        Run multiple agents in sequence.
        
        Args:
            agents: List of agent names to run in order
            model_type: AI model type (overrides config)
            language: Language for prompts (overrides detected)
            use_heuristic: Use heuristic mode (overrides config)
            use_multithreading: Use multi-threading (overrides config)
            max_workers: Max parallel workers (overrides config)
        
        Returns:
            Dictionary of final results by mapping key
        """
        if self.dfs is None or self.mapping_keys is None:
            raise ValueError("Data not loaded. Call load_data() first.")
        
        print("\n" + "="*60)
        print(f"RUNNING SEQUENTIAL AGENTS: {' -> '.join(agents)}")
        print("="*60 + "\n")
        
        results = None
        
        for i, agent_name in enumerate(agents):
            print(f"\n[{i+1}/{len(agents)}] Running {agent_name}...")
            
            results = self.run_agent_only(
                agent_name=agent_name,
                previous_results=results,
                model_type=model_type,
                language=language,
                use_heuristic=use_heuristic,
                use_multithreading=use_multithreading,
                max_workers=max_workers
            )
        
        self.results = results
        
        print("\n" + "="*60)
        print("ALL SEQUENTIAL AGENTS COMPLETED")
        print("="*60 + "\n")
        
        return results
    
    def save_results(self, output_path: str = 'fdd_utils/output/results.yml'):
        """
        Save results to YAML file.
        
        Args:
            output_path: Path to output file
        """
        if not self.results:
            print("No results to save.")
            return
        
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        
        with open(output_path, 'w', encoding='utf-8') as f:
            yaml.dump(self.results, f, default_flow_style=False, allow_unicode=True)
        
        print(f"Results saved to: {output_path}")
    
    def get_results(self) -> Dict[str, str]:
        """Get the current results."""
        return self.results
    
    def get_result_for_key(self, mapping_key: str) -> Optional[str]:
        """Get result for a specific mapping key."""
        return self.results.get(mapping_key)
    
    def print_results_summary(self):
        """Print a summary of results."""
        if not self.results:
            print("No results available.")
            return
        
        print("\n" + "="*60)
        print("RESULTS SUMMARY")
        print("="*60)
        
        for key, content in self.results.items():
            if content:
                preview = content[:100] + "..." if len(content) > 100 else content
                print(f"\n{key}:")
                print(f"  Length: {len(content)} chars")
                print(f"  Preview: {preview}")
            else:
                print(f"\n{key}: [FAILED]")
        
        print("\n" + "="*60 + "\n")


# Convenience function for quick pipeline execution
def run_quick_pipeline(
    databook_path: str,
    entity_name: str,
    mode: str = "All",
    model_type: str = "deepseek",
    language: str = "Eng",
    config_path: str = 'fdd_utils/config.yml'
) -> Dict[str, str]:
    """
    Quick execution of full AI pipeline.
    
    Args:
        databook_path: Path to Excel databook
        entity_name: Name of entity to extract
        mode: Filter mode for worksheets
        model_type: AI model type
        language: Language for prompts
        config_path: Path to config file
    
    Returns:
        Dictionary of final results by mapping key
    """
    orchestrator = AIPipelineOrchestrator(config_path=config_path)
    orchestrator.load_data(databook_path, entity_name, mode)
    results = orchestrator.run_full_pipeline(model_type=model_type, language=language)
    orchestrator.save_results()
    
    return results


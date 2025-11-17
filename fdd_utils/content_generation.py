"""
Enhanced Content Generation Module with Multi-Agent AI Pipeline
Supports: Content Generation, Value Checking, Content Refinement, and Format Checking
Features: Multi-threading, unified logging, tqdm progress bars
"""

import yaml
import os
import pandas as pd
import time
import re
import logging
from typing import Dict, List, Tuple, Any, Optional
from tqdm import tqdm
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from fdd_utils.ai_helper import AIHelper


class UnifiedLogger:
    """Unified logger for entire AI processing run."""
    
    def __init__(self, log_dir: str = 'fdd_utils/logs'):
        self.log_dir = log_dir
        os.makedirs(log_dir, exist_ok=True)
        
        # Create timestamp for this run
        self.run_id = datetime.now().strftime("%Y%m%d_%H%M%S")
        self.log_file = os.path.join(log_dir, f'ai_processing_{self.run_id}.log')
        self.log_data_file = os.path.join(log_dir, f'ai_data_{self.run_id}.yml')
        
        # Setup file and console logging
        self.logger = logging.getLogger(f'ContentGeneration_{self.run_id}')
        self.logger.setLevel(logging.INFO)
        self.logger.handlers = []  # Clear existing handlers
        
        # File handler
        fh = logging.FileHandler(self.log_file, encoding='utf-8')
        fh.setLevel(logging.DEBUG)
        
        # Console handler
        ch = logging.StreamHandler()
        ch.setLevel(logging.INFO)
        
        # Formatter
        formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
        fh.setFormatter(formatter)
        ch.setFormatter(formatter)
        
        self.logger.addHandler(fh)
        self.logger.addHandler(ch)
        
        # Data storage for YAML output
        self.run_data = {
            'run_id': self.run_id,
            'start_time': datetime.now().isoformat(),
            'agents_executed': [],
            'processing_results': {}
        }
        
        self.logger.info(f"=== Started new AI processing run: {self.run_id} ===")
    
    def log_agent_start(self, agent_name: str, mapping_key: str):
        """Log when an agent starts processing."""
        self.logger.info(f"[{agent_name}] Processing: {mapping_key}")
    
    def log_agent_complete(self, agent_name: str, mapping_key: str, result: Dict[str, Any], 
                          system_prompt: str = '', user_prompt: str = ''):
        """Log when an agent completes processing."""
        duration = result.get('duration', 0)
        tokens = result.get('tokens_used', 0)
        self.logger.info(
            f"[{agent_name}] Completed: {mapping_key} | "
            f"Duration: {duration:.2f}s | Tokens: {tokens}"
        )
        
        # Store in run data
        if mapping_key not in self.run_data['processing_results']:
            self.run_data['processing_results'][mapping_key] = {}
        
        self.run_data['processing_results'][mapping_key][agent_name] = {
            'duration': duration,
            'tokens_used': tokens,
            'mode': result.get('mode', 'ai'),
            'timestamp': datetime.now().isoformat(),
            'system_prompt': system_prompt[:500] if system_prompt else '',  # First 500 chars
            'user_prompt': user_prompt[:1000] if user_prompt else '',  # First 1000 chars
            'output': result.get('content', '')  # Full output
        }
    
    def log_error(self, agent_name: str, mapping_key: str, error: Exception):
        """Log errors during processing."""
        self.logger.error(f"[{agent_name}] Error processing {mapping_key}: {str(error)}")
        
        if mapping_key not in self.run_data['processing_results']:
            self.run_data['processing_results'][mapping_key] = {}
        
        self.run_data['processing_results'][mapping_key][agent_name] = {
            'error': str(error),
            'timestamp': datetime.now().isoformat()
        }
    
    def finalize(self):
        """Finalize logging and save data."""
        self.run_data['end_time'] = datetime.now().isoformat()
        
        # Calculate summary statistics
        total_duration = 0
        total_tokens = 0
        total_items = len(self.run_data['processing_results'])
        
        for key, agents in self.run_data['processing_results'].items():
            for agent, data in agents.items():
                if 'duration' in data:
                    total_duration += data['duration']
                if 'tokens_used' in data:
                    total_tokens += data['tokens_used']
        
        self.run_data['summary'] = {
            'total_items_processed': total_items,
            'total_duration_seconds': total_duration,
            'total_tokens_used': total_tokens,
            'agents_used': list(set(
                agent for agents in self.run_data['processing_results'].values() 
                for agent in agents.keys()
            ))
        }
        
        # Save to YAML
        with open(self.log_data_file, 'w', encoding='utf-8') as f:
            yaml.dump(self.run_data, f, default_flow_style=False, allow_unicode=True)
        
        self.logger.info(f"=== Completed AI processing run: {self.run_id} ===")
        self.logger.info(f"Summary: {total_items} items, {total_duration:.2f}s, {total_tokens} tokens")
        self.logger.info(f"Log files: {self.log_file}, {self.log_data_file}")


def map_value_to_component(value: str, component: Optional[str] = None, 
                          file_path: str = 'fdd_utils/mappings.yml') -> Any:
    """Map worksheet value to component from mappings."""
    with open(file_path, 'r', encoding='utf-8') as file:
        mappings = yaml.safe_load(file)
    
    for key, data in mappings.items():
        if value in data.get('aliases', []):
            if component and component in data:
                return data[component]
            return key
    return None


def load_prompts_and_format(
    agent_name: str, 
    language: str, 
    mapping_key: str, 
    df: pd.DataFrame,
    prompts_file: str = 'fdd_utils/prompts.yml',
    **kwargs
) -> Tuple[str, str]:
    """
    Load and format prompts for specified agent and language.
    
    Args:
        agent_name: Name of the agent (agent_1, agent_2, etc.)
        language: Language code ('Eng' or 'Chi')
        mapping_key: Mapping key for data
        df: DataFrame with financial data
        prompts_file: Path to prompts file
        **kwargs: Additional parameters for prompt formatting
    
    Returns:
        Tuple of (system_prompt, formatted_user_prompt)
    """
    with open(prompts_file, 'r', encoding='utf-8') as file:
        prompts_data = yaml.safe_load(file)
    
    agent_data = prompts_data.get(agent_name, {}).get(language, {})
    system_prompt = agent_data.get('system_prompt', '')
    user_prompt_template = agent_data.get('user_prompt', '')
    
    # Prepare formatting parameters
    format_params = {
        'key': mapping_key,
        **kwargs
    }
    
    # Add financial figure if df provided
    if df is not None and not df.empty:
        financial_figure_md = df.to_markdown(index=False).strip()
        format_params['financial_figure'] = financial_figure_md
        format_params['financial_data'] = financial_figure_md
    
    # Add patterns for agent_1
    if agent_name == 'agent_1':
        patterns = map_value_to_component(mapping_key, component='patterns')
        format_params['patterns'] = patterns
    
    try:
        user_prompt = user_prompt_template.format(**format_params)
    except KeyError as e:
        # If key is missing, return template as-is
        user_prompt = user_prompt_template
    
    return system_prompt, user_prompt


def process_single_item_agent_1(
    mapping_key: str,
    df: pd.DataFrame,
    model_type: str,
    language: str,
    use_heuristic: bool,
    logger: UnifiedLogger
) -> Tuple[str, str]:
    """Process single item through Agent 1 (Content Generator)."""
    try:
        logger.log_agent_start('agent_1', mapping_key)
        
        # Load and format prompts
        system_prompt, user_prompt = load_prompts_and_format(
            'agent_1', language, mapping_key, df
        )
        
        # Initialize AI Helper
        ai_helper = AIHelper(
            model_type=model_type,
            agent_name='agent_1',
            language=language,
            use_heuristic=use_heuristic
        )
        
        # Get response
        response = ai_helper.get_response(user_prompt, system_prompt)
        content = response['content'].strip().replace("\n\n", "\n").replace("\n \n", "\n")
        
        logger.log_agent_complete('agent_1', mapping_key, response, system_prompt, user_prompt)
        
        return mapping_key, content
        
    except Exception as e:
        logger.log_error('agent_1', mapping_key, e)
        return mapping_key, None


def process_single_item_agent_2(
    mapping_key: str,
    agent_1_output: str,
    df: pd.DataFrame,
    model_type: str,
    language: str,
    use_heuristic: bool,
    logger: UnifiedLogger
) -> Tuple[str, str]:
    """Process single item through Agent 2 (Value Checker)."""
    try:
        logger.log_agent_start('agent_2', mapping_key)
        
        # Get account type
        account_type = map_value_to_component(mapping_key, component='type')
        
        # Load and format prompts
        system_prompt, user_prompt = load_prompts_and_format(
            'agent_2', language, mapping_key, df,
            account=account_type,
            output=agent_1_output
        )
        
        # Initialize AI Helper
        ai_helper = AIHelper(
            model_type=model_type,
            agent_name='agent_2',
            language=language,
            use_heuristic=use_heuristic
        )
        
        # Get response
        response = ai_helper.get_response(user_prompt, system_prompt)
        content = response['content'].strip().replace("\n\n", "\n").replace("\n \n", "\n")
        
        logger.log_agent_complete('agent_2', mapping_key, response, system_prompt, user_prompt)
        
        return mapping_key, content
        
    except Exception as e:
        logger.log_error('agent_2', mapping_key, e)
        return mapping_key, agent_1_output  # Return previous output if error


def process_single_item_agent_3(
    mapping_key: str,
    agent_2_output: str,
    df: pd.DataFrame,
    model_type: str,
    language: str,
    use_heuristic: bool,
    logger: UnifiedLogger
) -> Tuple[str, str]:
    """Process single item through Agent 3 (Content Refiner)."""
    try:
        logger.log_agent_start('agent_3', mapping_key)
        
        original_length = len(agent_2_output)
        
        # Load and format prompts
        system_prompt, user_prompt = load_prompts_and_format(
            'agent_3', language, mapping_key, df,
            previous_content=agent_2_output,
            original_length=original_length
        )
        
        # Initialize AI Helper
        ai_helper = AIHelper(
            model_type=model_type,
            agent_name='agent_3',
            language=language,
            use_heuristic=use_heuristic
        )
        
        # Get response
        response = ai_helper.get_response(user_prompt, system_prompt)
        content = response['content'].strip().replace("\n\n", "\n").replace("\n \n", "\n")
        
        logger.log_agent_complete('agent_3', mapping_key, response, system_prompt, user_prompt)
        
        return mapping_key, content
        
    except Exception as e:
        logger.log_error('agent_3', mapping_key, e)
        return mapping_key, agent_2_output  # Return previous output if error


def process_single_item_agent_4(
    mapping_key: str,
    agent_3_output: str,
    model_type: str,
    language: str,
    use_heuristic: bool,
    logger: UnifiedLogger
) -> Tuple[str, str]:
    """Process single item through Agent 4 (Format Checker)."""
    try:
        logger.log_agent_start('agent_4', mapping_key)
        
        # Load and format prompts
        system_prompt, user_prompt = load_prompts_and_format(
            'agent_4', language, mapping_key, None,
            content=agent_3_output
        )
        
        # Initialize AI Helper
        ai_helper = AIHelper(
            model_type=model_type,
            agent_name='agent_4',
            language=language,
            use_heuristic=use_heuristic
        )
        
        # Get response
        response = ai_helper.get_response(user_prompt, system_prompt)
        content = response['content'].strip().replace("\n\n", "\n").replace("\n \n", "\n")
        
        logger.log_agent_complete('agent_4', mapping_key, response, system_prompt, user_prompt)
        
        return mapping_key, content
        
    except Exception as e:
        logger.log_error('agent_4', mapping_key, e)
        return mapping_key, agent_3_output  # Return previous output if error


def ai_pipeline_sequential_by_agent(
    mapping_keys: List[str],
    dfs: Dict[str, pd.DataFrame],
    model_type: str = 'deepseek',
    language: str = 'Eng',
    use_heuristic: bool = False,
    use_multithreading: bool = True,
    max_workers: int = 4
) -> Dict[str, str]:
    """
    Sequential-by-agent pipeline: Process ALL items through Agent 1, 
    then ALL items through Agent 2, etc. Multi-threading within each agent.
    
    Args:
        mapping_keys: List of mapping keys to process
        dfs: Dictionary of DataFrames by mapping key
        model_type: AI model type to use
        language: Language for prompts ('Eng' or 'Chi')
        use_heuristic: Whether to use heuristic mode
        use_multithreading: Whether to use multi-threading
        max_workers: Maximum number of parallel threads
    
    Returns:
        Dictionary of final results by mapping key
    """
    # Initialize unified logger
    logger = UnifiedLogger()
    logger.logger.info(f"Starting sequential-by-agent AI pipeline with {len(mapping_keys)} items")
    logger.logger.info(f"Model: {model_type}, Language: {language}, Multithreading: {use_multithreading}")
    
    # Storage for results at each stage
    results_agent_1 = {}
    results_agent_2 = {}
    results_agent_3 = {}
    results_agent_4 = {}
    
    # Agent 1: Process ALL items
    print("\n" + "="*60)
    print("AGENT 1: Content Generation")
    print("="*60)
    
    if use_multithreading and len(mapping_keys) > 1:
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            futures = {}
            for key in mapping_keys:
                if key in dfs:
                    future = executor.submit(
                        process_single_item_agent_1,
                        key, dfs[key], model_type, language, use_heuristic, logger
                    )
                    futures[future] = key
            
            with tqdm(total=len(futures), desc="Agent 1", unit='item') as pbar:
                for future in as_completed(futures):
                    mapping_key, content = future.result()
                    results_agent_1[mapping_key] = content
                    pbar.update(1)
    else:
        with tqdm(total=len(mapping_keys), desc="Agent 1", unit='item') as pbar:
            for key in mapping_keys:
                if key in dfs:
                    _, content = process_single_item_agent_1(
                        key, dfs[key], model_type, language, use_heuristic, logger
                    )
                    results_agent_1[key] = content
                pbar.update(1)
    
    # Agent 2: Process ALL items with Agent 1 outputs
    print("\n" + "="*60)
    print("AGENT 2: Value Checking")
    print("="*60)
    
    if use_multithreading and len(mapping_keys) > 1:
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            futures = {}
            for key in mapping_keys:
                if key in results_agent_1 and results_agent_1[key]:
                    future = executor.submit(
                        process_single_item_agent_2,
                        key, results_agent_1[key], dfs.get(key), 
                        model_type, language, use_heuristic, logger
                    )
                    futures[future] = key
            
            with tqdm(total=len(futures), desc="Agent 2", unit='item') as pbar:
                for future in as_completed(futures):
                    mapping_key, content = future.result()
                    results_agent_2[mapping_key] = content
                    pbar.update(1)
    else:
        with tqdm(total=len(mapping_keys), desc="Agent 2", unit='item') as pbar:
            for key in mapping_keys:
                if key in results_agent_1 and results_agent_1[key]:
                    _, content = process_single_item_agent_2(
                        key, results_agent_1[key], dfs.get(key),
                        model_type, language, use_heuristic, logger
                    )
                    results_agent_2[key] = content
                pbar.update(1)
    
    # Agent 3: Process ALL items with Agent 2 outputs
    print("\n" + "="*60)
    print("AGENT 3: Content Refinement")
    print("="*60)
    
    if use_multithreading and len(mapping_keys) > 1:
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            futures = {}
            for key in mapping_keys:
                if key in results_agent_2 and results_agent_2[key]:
                    future = executor.submit(
                        process_single_item_agent_3,
                        key, results_agent_2[key], dfs.get(key),
                        model_type, language, use_heuristic, logger
                    )
                    futures[future] = key
            
            with tqdm(total=len(futures), desc="Agent 3", unit='item') as pbar:
                for future in as_completed(futures):
                    mapping_key, content = future.result()
                    results_agent_3[mapping_key] = content
                    pbar.update(1)
    else:
        with tqdm(total=len(mapping_keys), desc="Agent 3", unit='item') as pbar:
            for key in mapping_keys:
                if key in results_agent_2 and results_agent_2[key]:
                    _, content = process_single_item_agent_3(
                        key, results_agent_2[key], dfs.get(key),
                        model_type, language, use_heuristic, logger
                    )
                    results_agent_3[key] = content
                pbar.update(1)
    
    # Agent 4: Process ALL items with Agent 3 outputs
    print("\n" + "="*60)
    print("AGENT 4: Format Checking")
    print("="*60)
    
    if use_multithreading and len(mapping_keys) > 1:
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            futures = {}
            for key in mapping_keys:
                if key in results_agent_3 and results_agent_3[key]:
                    future = executor.submit(
                        process_single_item_agent_4,
                        key, results_agent_3[key],
                        model_type, language, use_heuristic, logger
                    )
                    futures[future] = key
            
            with tqdm(total=len(futures), desc="Agent 4", unit='item') as pbar:
                for future in as_completed(futures):
                    mapping_key, content = future.result()
                    results_agent_4[mapping_key] = content
                    pbar.update(1)
    else:
        with tqdm(total=len(mapping_keys), desc="Agent 4", unit='item') as pbar:
            for key in mapping_keys:
                if key in results_agent_3 and results_agent_3[key]:
                    _, content = process_single_item_agent_4(
                        key, results_agent_3[key],
                        model_type, language, use_heuristic, logger
                    )
                    results_agent_4[key] = content
                pbar.update(1)
    
    # Finalize logging
    logger.finalize()
    
    print("\n" + "="*60)
    print("PIPELINE COMPLETED")
    print("="*60)
    print(f"Processed: {len(results_agent_4)} items")
    print(f"Successful: {sum(1 for v in results_agent_4.values() if v is not None)}")
    print(f"Failed: {sum(1 for v in results_agent_4.values() if v is None)}")
    print("="*60 + "\n")
    
    return results_agent_4


def ai_pipeline_full(
    mapping_keys: List[str],
    dfs: Dict[str, pd.DataFrame],
    model_type: str = 'deepseek',
    language: str = 'Eng',
    use_heuristic: bool = False,
    use_multithreading: bool = True,
    max_workers: int = 4
) -> Dict[str, str]:
    """
    Full AI pipeline with all 4 agents processing sequentially per item,
    but items can be processed in parallel.
    
    Args:
        mapping_keys: List of mapping keys to process
        dfs: Dictionary of DataFrames by mapping key
        model_type: AI model type to use
        language: Language for prompts ('Eng' or 'Chi')
        use_heuristic: Whether to use heuristic mode
        use_multithreading: Whether to use multi-threading
        max_workers: Maximum number of parallel threads
    
    Returns:
        Dictionary of final results by mapping key
    """
    # Initialize unified logger
    logger = UnifiedLogger()
    logger.logger.info(f"Starting full AI pipeline with {len(mapping_keys)} items")
    logger.logger.info(f"Model: {model_type}, Language: {language}, Multithreading: {use_multithreading}")
    
    results = {}
    
    def process_item_full_pipeline(mapping_key: str) -> Tuple[str, str]:
        """Process single item through all 4 agents."""
        if mapping_key not in dfs:
            return mapping_key, None
        
        df = dfs[mapping_key]
        
        # Agent 1: Content Generation
        _, content_1 = process_single_item_agent_1(
            mapping_key, df, model_type, language, use_heuristic, logger
        )
        
        if content_1 is None:
            return mapping_key, None
        
        # Agent 2: Value Check
        _, content_2 = process_single_item_agent_2(
            mapping_key, content_1, df, model_type, language, use_heuristic, logger
        )
        
        # Agent 3: Content Refinement
        _, content_3 = process_single_item_agent_3(
            mapping_key, content_2, df, model_type, language, use_heuristic, logger
        )
        
        # Agent 4: Format Check
        _, final_content = process_single_item_agent_4(
            mapping_key, content_3, model_type, language, use_heuristic, logger
        )
        
        return mapping_key, final_content
    
    # Process with or without multithreading
    if use_multithreading and len(mapping_keys) > 1:
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            futures = {
                executor.submit(process_item_full_pipeline, key): key 
                for key in mapping_keys
            }
            
            with tqdm(total=len(mapping_keys), desc="AI Pipeline Progress", unit='item') as pbar:
                for future in as_completed(futures):
                    mapping_key, final_content = future.result()
                    results[mapping_key] = final_content
                    pbar.update(1)
    else:
        # Sequential processing
        with tqdm(total=len(mapping_keys), desc="AI Pipeline Progress", unit='item') as pbar:
            for mapping_key in mapping_keys:
                _, final_content = process_item_full_pipeline(mapping_key)
                results[mapping_key] = final_content
                pbar.update(1)
    
    # Finalize logging
    logger.finalize()
    
    return results


def ai_pipeline_agent_only(
    agent_name: str,
    mapping_keys: List[str],
    dfs: Dict[str, pd.DataFrame],
    model_type: str = 'deepseek',
    language: str = 'Eng',
    use_heuristic: bool = False,
    use_multithreading: bool = True,
    max_workers: int = 4,
    previous_results: Optional[Dict[str, str]] = None
) -> Dict[str, str]:
    """
    Run only a specific agent in the pipeline.
    
    Args:
        agent_name: Name of agent to run ('agent_1', 'agent_2', 'agent_3', 'agent_4')
        mapping_keys: List of mapping keys to process
        dfs: Dictionary of DataFrames by mapping key
        model_type: AI model type to use
        language: Language for prompts ('Eng' or 'Chi')
        use_heuristic: Whether to use heuristic mode
        use_multithreading: Whether to use multi-threading
        max_workers: Maximum number of parallel threads
        previous_results: Previous results (required for agent_2, agent_3, agent_4)
    
    Returns:
        Dictionary of results by mapping key
    """
    logger = UnifiedLogger()
    logger.logger.info(f"Running {agent_name} only with {len(mapping_keys)} items")
    
    results = {}
    
    # Select appropriate processing function
    if agent_name == 'agent_1':
        process_func = lambda key: process_single_item_agent_1(
            key, dfs.get(key), model_type, language, use_heuristic, logger
        )
    elif agent_name == 'agent_2' and previous_results:
        process_func = lambda key: process_single_item_agent_2(
            key, previous_results.get(key, ''), dfs.get(key), 
            model_type, language, use_heuristic, logger
        )
    elif agent_name == 'agent_3' and previous_results:
        process_func = lambda key: process_single_item_agent_3(
            key, previous_results.get(key, ''), dfs.get(key),
            model_type, language, use_heuristic, logger
        )
    elif agent_name == 'agent_4' and previous_results:
        process_func = lambda key: process_single_item_agent_4(
            key, previous_results.get(key, ''),
            model_type, language, use_heuristic, logger
        )
    else:
        logger.logger.error(f"Invalid agent name or missing previous results: {agent_name}")
        logger.finalize()
        return {}
    
    # Process with or without multithreading
    if use_multithreading and len(mapping_keys) > 1:
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            futures = {executor.submit(process_func, key): key for key in mapping_keys}
            
            with tqdm(total=len(mapping_keys), desc=f"{agent_name} Progress", unit='item') as pbar:
                for future in as_completed(futures):
                    mapping_key, content = future.result()
                    results[mapping_key] = content
                    pbar.update(1)
    else:
        with tqdm(total=len(mapping_keys), desc=f"{agent_name} Progress", unit='item') as pbar:
            for mapping_key in mapping_keys:
                _, content = process_func(mapping_key)
                results[mapping_key] = content
                pbar.update(1)
    
    logger.finalize()
    return results

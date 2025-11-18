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


def clean_agent_output(content: str) -> str:
    """
    Clean agent output by removing common meta-commentary patterns.
    
    Args:
        content: Raw content from agent
        
    Returns:
        Cleaned content without meta-commentary
    """
    # Remove common meta-commentary prefixes (case-insensitive)
    # Note: Don't use (?i) inline flags with flags=re.IGNORECASE parameter
    prefixes_to_remove = [
        r'^verified\s+output:\s*',
        r'^corrected\s+output:\s*',
        r'^refined\s+output:\s*',
        r'^formatted\s+output:\s*',
        r'^final\s+output:\s*',
        r'^after\s+verification[,:]?\s*',
        r'^after\s+refining[,:]?\s*',
        r'^final\s+formatted\s+content:\s*',
        r'^the\s+corrected\s+output\s+is:\s*',
        r'^here\s+is\s+the\s+(corrected|refined|verified)\s+output:\s*',
        # Chinese patterns
        r'^已验证输出：\s*',
        r'^已更正输出：\s*',
        r'^精炼后的输出：\s*',
        r'^格式化后的输出：\s*',
        r'^经过验证[，,]\s*',
        r'^经过精炼后[，,]\s*',
    ]
    
    cleaned = content.strip()
    
    # Try to remove prefixes (case-insensitive)
    for pattern in prefixes_to_remove:
        cleaned = re.sub(pattern, '', cleaned, flags=re.IGNORECASE)
    
    # Remove surrounding quotes if they wrap the entire content
    if (cleaned.startswith('"') and cleaned.endswith('"')) or \
       (cleaned.startswith("'") and cleaned.endswith("'")):
        cleaned = cleaned[1:-1]
    
    # Remove meta-commentary at the end (sentences that mention verification/corrections)
    end_patterns = [
        r'\s*I\s+(?:verified|corrected|refined|checked).*$',
        r'\s*(?:Corrections?|Verifications?)\s+made:.*$',
        r'\s*我(?:验证|更正|精炼|检查)了.*$',
        r'\s*所做更正：.*$',
    ]
    
    for pattern in end_patterns:
        cleaned = re.sub(pattern, '', cleaned, flags=re.IGNORECASE)
    
    return cleaned.strip()


class UnifiedLogger:
    """Unified logger for entire AI processing run."""
    
    def __init__(self, log_dir: str = 'fdd_utils/logs', output_dir: str = 'fdd_utils/output'):
        self.log_dir = log_dir
        self.output_dir = output_dir
        os.makedirs(log_dir, exist_ok=True)
        os.makedirs(output_dir, exist_ok=True)
        
        # Create timestamp for this run
        self.run_id = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        # Create subfolder for this run
        self.run_folder = os.path.join(log_dir, f'run_{self.run_id}')
        os.makedirs(self.run_folder, exist_ok=True)
        
        # Store log and data files in the same subfolder
        self.log_file = os.path.join(self.run_folder, 'processing.log')
        self.log_data_file = os.path.join(self.run_folder, 'data.yml')
        self.results_file = os.path.join(self.run_folder, 'results.yml')
        
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
        content = result.get('content', '')
        
        # Console log with progress
        self.logger.info(
            f"[{agent_name}] ✅ {mapping_key} | "
            f"Duration: {duration:.2f}s | Tokens: {tokens} | "
            f"Content: {len(content)} chars"
        )
        
        # Store in run data with full details
        if mapping_key not in self.run_data['processing_results']:
            self.run_data['processing_results'][mapping_key] = {}
        
        self.run_data['processing_results'][mapping_key][agent_name] = {
            'duration': duration,
            'tokens_used': tokens,
            'mode': result.get('mode', 'ai'),
            'timestamp': datetime.now().isoformat(),
            'system_prompt': system_prompt[:500] if system_prompt else '',
            'user_prompt': user_prompt[:1000] if user_prompt else '',
            'output': content,  # Full output
            'content_length': len(content)
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
    
    def finalize(self, results: Dict[str, Dict[str, str]] = None):
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
        
        # Save results to the same subfolder if provided
        if results:
            with open(self.results_file, 'w', encoding='utf-8') as f:
                yaml.dump(results, f, default_flow_style=False, allow_unicode=True)
        
        self.logger.info(f"=== Completed AI processing run: {self.run_id} ===")
        self.logger.info(f"Summary: {total_items} items, {total_duration:.2f}s, {total_tokens} tokens")
        self.logger.info(f"Run folder: {self.run_folder}")
        self.logger.info(f"Files saved: {self.log_file}, {self.log_data_file}, {self.results_file}")


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
    mappings_file: str = 'fdd_utils/mappings.yml',
    **kwargs
) -> Tuple[str, str]:
    """
    Load and format prompts for specified agent and language.
    
    For 1_Generator/agent_1: Loads from mappings.yml (account-specific prompts)
    For 2_Auditor/3_Refiner/4_Validator: Loads from prompts.yml (generic prompts)
    
    Args:
        agent_name: Name of the agent (agent_1 or 1_Generator, agent_2 or 2_Auditor, etc.)
        language: Language code ('Eng' or 'Chi')
        mapping_key: Mapping key for data
        df: DataFrame with financial data
        prompts_file: Path to prompts file (for agent_2-4)
        mappings_file: Path to mappings file (for agent_1)
        **kwargs: Additional parameters for prompt formatting
    
    Returns:
        Tuple of (system_prompt, formatted_user_prompt)
    """
    # Agent 1/1_Generator: Read from mappings.yml (account-specific)
    if agent_name in ['agent_1', '1_Generator']:
        with open(mappings_file, 'r', encoding='utf-8') as file:
            mappings_data = yaml.safe_load(file)
        
        account_data = mappings_data.get(mapping_key, {})
        agent_prompts = account_data.get('agent_1_prompts', {}).get(language, {})
        system_prompt = agent_prompts.get('system_prompt', '')
        user_prompt_template = agent_prompts.get('user_prompt', '')
        
        # Fallback to generic agent_1 prompts if account-specific not found
        if not system_prompt or not user_prompt_template:
            generic_prompts = mappings_data.get('_default_agent_1', {}).get(language, {})
            if not system_prompt:
                system_prompt = generic_prompts.get('system_prompt', '')
            if not user_prompt_template:
                user_prompt_template = generic_prompts.get('user_prompt', '')
    
    # Agent 2-4: Read from prompts.yml (generic)
    else:
        with open(prompts_file, 'r', encoding='utf-8') as file:
            prompts_data = yaml.safe_load(file)
        
        # Map old names to new names for backward compatibility
        agent_key = agent_name
        if agent_name == 'agent_2':
            agent_key = '2_Auditor'
        elif agent_name == 'agent_3':
            agent_key = '3_Refiner'
        elif agent_name == 'agent_4':
            agent_key = '4_Validator'
        
        agent_data = prompts_data.get(agent_key, {}).get(language, {})
        system_prompt = agent_data.get('system_prompt', '')
        user_prompt_template = agent_data.get('user_prompt', '')
    
    # Prepare formatting parameters
    format_params = {
        'key': mapping_key,
        'language': language,  # Add language parameter for agent_4
        **kwargs
    }
    
    # Add financial figure if df provided
    if df is not None and not df.empty:
        # Format DataFrame to avoid scientific notation
        df_formatted = df.copy()
        for col in df_formatted.columns:
            if pd.api.types.is_numeric_dtype(df_formatted[col]):
                # Format numbers to avoid scientific notation
                df_formatted[col] = df_formatted[col].apply(
                    lambda x: f"{x:,.2f}" if pd.notna(x) else x
                )
        
        financial_figure_md = df_formatted.to_markdown(index=False).strip()
        format_params['financial_figure'] = financial_figure_md
        format_params['financial_data'] = financial_figure_md
    
    # Add patterns for agent_1/1_Generator
    if agent_name in ['agent_1', '1_Generator']:
        patterns = map_value_to_component(mapping_key, component='patterns')
        format_params['patterns'] = patterns
    
    try:
        user_prompt = user_prompt_template.format(**format_params)
    except KeyError as e:
        # If key is missing, log warning and return template as-is
        import logging
        logger = logging.getLogger('load_prompts_and_format')
        logger.warning(f"Missing key in prompt template for {agent_name}: {e}")
        logger.warning(f"Available keys: {list(format_params.keys())}")
        user_prompt = user_prompt_template
    
    return system_prompt, user_prompt


def process_single_item_agent_1(
    mapping_key: str,
    df: pd.DataFrame,
    ai_helper: AIHelper,
    logger: UnifiedLogger
) -> Tuple[str, str]:
    """Process single item through Agent 1 (Content Generator)."""
    try:
        logger.log_agent_start('agent_1', mapping_key)
        
        # Load and format prompts
        system_prompt, user_prompt = load_prompts_and_format(
            'agent_1', ai_helper.language, mapping_key, df
        )
        
        # Get response (reuse AIHelper)
        response = ai_helper.get_response(user_prompt, system_prompt)
        content = response['content'].strip().replace("\n\n", "\n").replace("\n \n", "\n")
        
        # Clean agent output to remove meta-commentary
        content = clean_agent_output(content)
        
        logger.log_agent_complete('agent_1', mapping_key, response, system_prompt, user_prompt)
        
        return mapping_key, content
        
    except Exception as e:
        logger.log_error('agent_1', mapping_key, e)
        return mapping_key, None


def process_single_item_agent_2(
    mapping_key: str,
    agent_1_output: str,
    df: pd.DataFrame,
    ai_helper: AIHelper,
    logger: UnifiedLogger
) -> Tuple[str, str]:
    """Process single item through Agent 2 (Value Checker)."""
    try:
        logger.log_agent_start('agent_2', mapping_key)
        
        # Get account type
        account_type = map_value_to_component(mapping_key, component='type')
        
        # Load and format prompts
        system_prompt, user_prompt = load_prompts_and_format(
            'agent_2', ai_helper.language, mapping_key, df,
            account=account_type,
            output=agent_1_output
        )
        
        # Get response (reuse AIHelper)
        response = ai_helper.get_response(user_prompt, system_prompt)
        content = response['content'].strip().replace("\n\n", "\n").replace("\n \n", "\n")
        
        # Clean agent output to remove meta-commentary
        content = clean_agent_output(content)
        
        logger.log_agent_complete('agent_2', mapping_key, response, system_prompt, user_prompt)
        
        return mapping_key, content
        
    except Exception as e:
        logger.log_error('agent_2', mapping_key, e)
        return mapping_key, agent_1_output  # Return previous output if error


def process_single_item_agent_3(
    mapping_key: str,
    agent_2_output: str,
    df: pd.DataFrame,
    ai_helper: AIHelper,
    logger: UnifiedLogger
) -> Tuple[str, str]:
    """Process single item through Agent 3 (Content Refiner)."""
    try:
        logger.log_agent_start('agent_3', mapping_key)
        
        original_length = len(agent_2_output)
        
        # Load and format prompts
        system_prompt, user_prompt = load_prompts_and_format(
            'agent_3', ai_helper.language, mapping_key, df,
            previous_content=agent_2_output,
            original_length=original_length
        )
        
        # Get response (reuse AIHelper)
        response = ai_helper.get_response(user_prompt, system_prompt)
        content = response['content'].strip().replace("\n\n", "\n").replace("\n \n", "\n")
        
        # Clean agent output to remove meta-commentary
        content = clean_agent_output(content)
        
        logger.log_agent_complete('agent_3', mapping_key, response, system_prompt, user_prompt)
        
        return mapping_key, content
        
    except Exception as e:
        logger.log_error('agent_3', mapping_key, e)
        return mapping_key, agent_2_output  # Return previous output if error


def process_single_item_agent_4(
    mapping_key: str,
    agent_3_output: str,
    ai_helper: AIHelper,
    logger: UnifiedLogger
) -> Tuple[str, str]:
    """Process single item through Agent 4 (Format Checker)."""
    try:
        logger.log_agent_start('agent_4', mapping_key)
        
        # Load and format prompts
        system_prompt, user_prompt = load_prompts_and_format(
            'agent_4', ai_helper.language, mapping_key, None,
            content=agent_3_output
        )
        
        # Get response (reuse AIHelper)
        response = ai_helper.get_response(user_prompt, system_prompt)
        content = response['content'].strip().replace("\n\n", "\n").replace("\n \n", "\n")
        
        # Clean agent output to remove meta-commentary
        content = clean_agent_output(content)
        
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
    max_workers: Optional[int] = None
) -> Dict[str, Dict[str, str]]:
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
        max_workers: Maximum number of parallel threads (None = use all CPU cores)
    
    Returns:
        Dictionary with structure: {mapping_key: {agent_1: content, agent_2: content, ...}}
    """
    # Auto-detect number of workers if not specified
    if max_workers is None:
        import multiprocessing
        max_workers = multiprocessing.cpu_count()
    
    # Initialize unified logger
    logger = UnifiedLogger()
    logger.logger.info(f"Starting sequential-by-agent AI pipeline with {len(mapping_keys)} items")
    logger.logger.info(f"Model: {model_type}, Language: {language}, Multithreading: {use_multithreading}, Workers: {max_workers}")
    
    # Create ONE reusable AIHelper instance
    ai_helper = AIHelper(
        model_type=model_type,
        agent_name='content_pipeline',  # Generic name since we're reusing it
        language=language,
        use_heuristic=use_heuristic
    )
    
    # Storage for ALL results from all agents
    final_results = {key: {} for key in mapping_keys if key in dfs}
    
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
                        key, dfs[key], ai_helper, logger
                    )
                    futures[future] = key
            
            with tqdm(total=len(futures), desc="Agent 1", unit='item') as pbar:
                for future in as_completed(futures):
                    mapping_key, content = future.result()
                    if mapping_key in final_results:
                        final_results[mapping_key]['agent_1'] = content
                    pbar.update(1)
    else:
        with tqdm(total=len(mapping_keys), desc="Agent 1", unit='item') as pbar:
            for key in mapping_keys:
                if key in dfs:
                    _, content = process_single_item_agent_1(
                        key, dfs[key], ai_helper, logger
                    )
                    if key in final_results:
                        final_results[key]['agent_1'] = content
                pbar.update(1)
    
    # Agent 2: Process ALL items with Agent 1 outputs
    print("\n" + "="*60)
    print("AGENT 2: Value Checking")
    print("="*60)
    
    if use_multithreading and len(mapping_keys) > 1:
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            futures = {}
            for key in mapping_keys:
                if key in final_results and 'agent_1' in final_results[key]:
                    future = executor.submit(
                        process_single_item_agent_2,
                        key, final_results[key]['agent_1'], dfs.get(key), 
                        ai_helper, logger
                    )
                    futures[future] = key
            
            with tqdm(total=len(futures), desc="Agent 2", unit='item') as pbar:
                for future in as_completed(futures):
                    mapping_key, content = future.result()
                    if mapping_key in final_results:
                        final_results[mapping_key]['agent_2'] = content
                    pbar.update(1)
    else:
        with tqdm(total=len(mapping_keys), desc="Agent 2", unit='item') as pbar:
            for key in mapping_keys:
                if key in final_results and 'agent_1' in final_results[key]:
                    _, content = process_single_item_agent_2(
                        key, final_results[key]['agent_1'], dfs.get(key),
                        ai_helper, logger
                    )
                    if key in final_results:
                        final_results[key]['agent_2'] = content
                pbar.update(1)
    
    # Agent 3: Process ALL items with Agent 2 outputs
    print("\n" + "="*60)
    print("AGENT 3: Content Refinement")
    print("="*60)
    
    if use_multithreading and len(mapping_keys) > 1:
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            futures = {}
            for key in mapping_keys:
                if key in final_results and 'agent_2' in final_results[key]:
                    future = executor.submit(
                        process_single_item_agent_3,
                        key, final_results[key]['agent_2'], dfs.get(key),
                        ai_helper, logger
                    )
                    futures[future] = key
            
            with tqdm(total=len(futures), desc="Agent 3", unit='item') as pbar:
                for future in as_completed(futures):
                    mapping_key, content = future.result()
                    if mapping_key in final_results:
                        final_results[mapping_key]['agent_3'] = content
                    pbar.update(1)
    else:
        with tqdm(total=len(mapping_keys), desc="Agent 3", unit='item') as pbar:
            for key in mapping_keys:
                if key in final_results and 'agent_2' in final_results[key]:
                    _, content = process_single_item_agent_3(
                        key, final_results[key]['agent_2'], dfs.get(key),
                        ai_helper, logger
                    )
                    if key in final_results:
                        final_results[key]['agent_3'] = content
                pbar.update(1)
    
    # Agent 4: Process ALL items with Agent 3 outputs
    print("\n" + "="*60)
    print("AGENT 4: Format Checking")
    print("="*60)
    
    if use_multithreading and len(mapping_keys) > 1:
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            futures = {}
            for key in mapping_keys:
                if key in final_results and 'agent_3' in final_results[key]:
                    future = executor.submit(
                        process_single_item_agent_4,
                        key, final_results[key]['agent_3'],
                        ai_helper, logger
                    )
                    futures[future] = key
            
            with tqdm(total=len(futures), desc="Agent 4", unit='item') as pbar:
                for future in as_completed(futures):
                    mapping_key, content = future.result()
                    if mapping_key in final_results:
                        final_results[mapping_key]['agent_4'] = content
                        final_results[mapping_key]['final'] = content  # Also store as final
                    pbar.update(1)
    else:
        with tqdm(total=len(mapping_keys), desc="Agent 4", unit='item') as pbar:
            for key in mapping_keys:
                if key in final_results and 'agent_3' in final_results[key]:
                    _, content = process_single_item_agent_4(
                        key, final_results[key]['agent_3'],
                        ai_helper, logger
                    )
                    if key in final_results:
                        final_results[key]['agent_4'] = content
                        final_results[key]['final'] = content  # Also store as final
                pbar.update(1)
    
    # Finalize logging and save results
    logger.finalize(final_results)
    
    print("\n" + "="*60)
    print("PIPELINE COMPLETED")
    print("="*60)
    print(f"Processed: {len(final_results)} items")
    print(f"Successful: {sum(1 for v in final_results.values() if 'final' in v)}")
    print(f"Failed: {sum(1 for v in final_results.values() if 'final' not in v)}")
    print(f"Results saved to: {logger.run_folder}")
    print("="*60 + "\n")
    
    return final_results


def run_ai_pipeline(
    mapping_keys: List[str],
    dfs: Dict[str, pd.DataFrame],
    model_type: str = 'deepseek',
    language: str = 'Eng',
    use_heuristic: bool = False,
    use_multithreading: bool = True,
    max_workers: Optional[int] = None
) -> Dict[str, Dict[str, str]]:
    """
    Simple AI pipeline: Run all items through 4 agents sequentially.
    This is the main function to use.
    
    Args:
        mapping_keys: List of mapping keys to process
        dfs: Dictionary of DataFrames by mapping key
        model_type: AI model type ('openai', 'local', 'deepseek')
        language: Language for prompts ('Eng' or 'Chi')
        use_heuristic: Use rule-based processing instead of AI
        use_multithreading: Use multi-threading within each agent
        max_workers: Number of parallel workers (None = use all CPU cores)
    
    Returns:
        Dict with structure: {
            'Cash': {
                'agent_1': '...',
                'agent_2': '...',
                'agent_3': '...',
                'agent_4': '...',
                'final': '...'
            }, ...
        }
    """
    return ai_pipeline_sequential_by_agent(
        mapping_keys=mapping_keys,
        dfs=dfs,
        model_type=model_type,
        language=language,
        use_heuristic=use_heuristic,
        use_multithreading=use_multithreading,
        max_workers=max_workers
    )


def save_results(results: Dict[str, Dict[str, str]], output_path: str = 'fdd_utils/output/results.yml'):
    """
    Save pipeline results to YAML file.
    
    Args:
        results: Results dictionary from run_ai_pipeline()
        output_path: Path to save results
    """
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    
    with open(output_path, 'w', encoding='utf-8') as f:
        yaml.dump(results, f, default_flow_style=False, allow_unicode=True)
    
    print(f"✅ Results saved to: {output_path}")


def extract_final_contents(results: Dict[str, Dict[str, str]]) -> Dict[str, str]:
    """
    Extract just the final contents for feeding into patterns.
    
    Args:
        results: Results dictionary from run_ai_pipeline()
    
    Returns:
        Dict with {key: final_content}
    """
    final_contents = {}
    
    for key, value in results.items():
        if isinstance(value, dict) and 'final' in value:
            final_contents[key] = value['final']
    
    return final_contents

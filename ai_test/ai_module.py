"""
AI Module for Testing
Provides AI generation capabilities with support for multiple providers
"""

import json
import os
import sys
from pathlib import Path
from openai import OpenAI

# Add parent directory to path for imports
current_dir = Path(__file__).resolve().parent
parent_dir = current_dir.parent
if str(parent_dir) not in sys.path:
    sys.path.insert(0, str(parent_dir))


class AIModule:
    """AI Module for testing financial report generation"""
    
    def __init__(self, config_path=None):
        """Initialize AI Module with configuration
        
        Args:
            config_path: Path to config.json file. If None, looks in current directory
        """
        if config_path is None:
            config_path = current_dir / "config.json"
        
        self.config = self._load_config(config_path)
        self.clients = {}
        self._initialize_clients()
    
    def _load_config(self, config_path):
        """Load configuration from JSON file"""
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                config = json.load(f)
            print(f"✅ Configuration loaded from: {config_path}")
            return config
        except Exception as e:
            print(f"❌ Error loading config: {e}")
            return {}
    
    def _initialize_clients(self):
        """Initialize AI clients based on configuration"""
        try:
            # DeepSeek Client
            if self.config.get('DEEPSEEK_API_KEY'):
                self.clients['deepseek'] = OpenAI(
                    api_key=self.config['DEEPSEEK_API_KEY'],
                    base_url=self.config['DEEPSEEK_API_BASE']
                )
                print("✅ DeepSeek client initialized")
            
            # OpenAI Client
            if self.config.get('OPENAI_API_KEY') and 'placeholder' not in self.config['OPENAI_API_KEY'].lower():
                self.clients['openai'] = OpenAI(
                    api_key=self.config['OPENAI_API_KEY'],
                    base_url=self.config['OPENAI_API_BASE']
                )
                print("✅ OpenAI client initialized")
            
            # Local AI Client
            if self.config.get('LOCAL_AI_ENABLED'):
                self.clients['local'] = OpenAI(
                    api_key=self.config.get('LOCAL_AI_API_KEY', 'not-needed'),
                    base_url=self.config['LOCAL_AI_API_BASE']
                )
                print("✅ Local AI client initialized")
                
        except Exception as e:
            print(f"❌ Error initializing clients: {e}")
    
    def generate_content(
        self, 
        system_prompt, 
        user_prompt, 
        provider='deepseek',
        model=None,
        temperature=0.7,
        max_tokens=2000
    ):
        """Generate content using specified AI provider
        
        Args:
            system_prompt: System prompt for AI
            user_prompt: User prompt/query
            provider: AI provider ('deepseek', 'openai', 'local')
            model: Specific model to use (optional, uses default from config)
            temperature: Temperature for generation (0.0-2.0)
            max_tokens: Maximum tokens to generate
            
        Returns:
            dict with 'content', 'tokens', and 'model' keys
        """
        try:
            # Get client
            if provider not in self.clients:
                return {
                    'error': f"Provider '{provider}' not available",
                    'content': None
                }
            
            client = self.clients[provider]
            
            # Determine model
            if model is None:
                if provider == 'deepseek':
                    model = self.config.get('DEEPSEEK_CHAT_MODEL', 'deepseek-chat')
                elif provider == 'openai':
                    model = self.config.get('OPENAI_CHAT_MODEL', 'gpt-4o-mini')
                elif provider == 'local':
                    model = self.config.get('LOCAL_AI_CHAT_MODEL', 'local-qwen2')
            
            print(f"🤖 Generating content using {provider} ({model})...")
            
            # Make API call
            response = client.chat.completions.create(
                model=model,
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": user_prompt}
                ],
                temperature=temperature,
                max_tokens=max_tokens
            )
            
            # Extract response
            content = response.choices[0].message.content
            tokens = {
                'prompt_tokens': response.usage.prompt_tokens,
                'completion_tokens': response.usage.completion_tokens,
                'total_tokens': response.usage.total_tokens
            }
            
            print(f"✅ Content generated ({tokens['total_tokens']} tokens)")
            
            return {
                'content': content,
                'tokens': tokens,
                'model': model,
                'provider': provider
            }
            
        except Exception as e:
            print(f"❌ Error generating content: {e}")
            return {
                'error': str(e),
                'content': None
            }
    
    def test_agent1(
        self,
        financial_data,
        key,
        entity_name,
        mode='english',
        provider='deepseek'
    ):
        """Test Agent 1 (Content Generation) individually
        
        Args:
            financial_data: Financial table data (markdown format)
            key: Financial key (e.g., 'Cash', 'AR')
            entity_name: Entity name
            mode: Language mode ('english' or 'chinese')
            provider: AI provider to use
            
        Returns:
            dict with agent1 result
        """
        print(f"\n{'='*80}")
        print(f"🤖 TESTING AGENT 1 ONLY (Content Generation)")
        print(f"   Key: {key}")
        print(f"   Entity: {entity_name}")
        print(f"   Mode: {mode}")
        print(f"   Provider: {provider}")
        print(f"{'='*80}\n")
        
        # Load prompts
        prompts_path = parent_dir / "fdd_utils" / "prompts.json"
        try:
            with open(prompts_path, 'r', encoding='utf-8') as f:
                prompts = json.load(f)
        except Exception as e:
            print(f"❌ Error loading prompts: {e}")
            return {'error': 'Could not load prompts'}
        
        # Get system prompts for mode
        system_prompts = prompts.get('system_prompts', {}).get(mode, {})
        
        # Agent 1: Content Generation
        agent1_system = system_prompts.get('Agent 1', '')
        agent1_user = f"""
Generate financial commentary for {key} - {entity_name}

FINANCIAL DATA:
{financial_data}

Instructions:
1. Analyze the financial data provided
2. Generate professional commentary following the pattern
3. Include specific figures and entities from the data
4. Keep it concise (100-120 words maximum)
5. Do NOT output raw table data
"""
        
        result = self.generate_content(
            system_prompt=agent1_system,
            user_prompt=agent1_user,
            provider=provider
        )
        
        if result.get('content'):
            print(f"\n📝 Agent 1 Output ({len(result['content'])} chars):")
            print("-" * 80)
            print(result['content'])
            print("-" * 80)
            print(f"✅ Tokens Used: {result['tokens']['total_tokens']}")
        
        return result
    
    def test_agent2(
        self,
        financial_data,
        agent1_content,
        mode='english',
        provider='deepseek'
    ):
        """Test Agent 2 (Proofreader) individually
        
        Args:
            financial_data: Original financial table data
            agent1_content: Content generated by Agent 1
            mode: Language mode ('english' or 'chinese')
            provider: AI provider to use
            
        Returns:
            dict with agent2 result
        """
        print(f"\n{'='*80}")
        print(f"🔍 TESTING AGENT 2 ONLY (AI Proofreader)")
        print(f"   Mode: {mode}")
        print(f"   Provider: {provider}")
        print(f"{'='*80}\n")
        
        # Load prompts
        prompts_path = parent_dir / "fdd_utils" / "prompts.json"
        try:
            with open(prompts_path, 'r', encoding='utf-8') as f:
                prompts = json.load(f)
        except Exception as e:
            print(f"❌ Error loading prompts: {e}")
            return {'error': 'Could not load prompts'}
        
        # Get system prompts for mode
        system_prompts = prompts.get('system_prompts', {}).get(mode, {})
        
        agent2_system = system_prompts.get('AI Proofreader', '')
        agent2_user = f"""
Review the following content for compliance:

ORIGINAL FINANCIAL DATA:
{financial_data}

AGENT 1 CONTENT:
{agent1_content}

Provide your review in JSON format with:
- is_compliant: boolean
- issues: array of issues found
- corrected_content: the corrected version
- figure_checks: array of figure validation results
- entity_checks: array of entity validation results
- grammar_notes: array of grammar improvements
"""
        
        result = self.generate_content(
            system_prompt=agent2_system,
            user_prompt=agent2_user,
            provider=provider,
            max_tokens=3000
        )
        
        if result.get('content'):
            print(f"\n✅ Agent 2 Output ({len(result['content'])} chars):")
            print("-" * 80)
            print(result['content'][:500] + "..." if len(result['content']) > 500 else result['content'])
            print("-" * 80)
            print(f"✅ Tokens Used: {result['tokens']['total_tokens']}")
        
        return result
    
    def test_agent3(
        self,
        agent1_content,
        mode='english',
        provider='deepseek'
    ):
        """Test Agent 3 (Pattern Compliance Validator) individually
        
        Args:
            agent1_content: Content generated by Agent 1
            mode: Language mode ('english' or 'chinese')
            provider: AI provider to use
            
        Returns:
            dict with agent3 result
        """
        print(f"\n{'='*80}")
        print(f"✨ TESTING AGENT 3 ONLY (Pattern Compliance)")
        print(f"   Mode: {mode}")
        print(f"   Provider: {provider}")
        print(f"{'='*80}\n")
        
        # Load prompts
        prompts_path = parent_dir / "fdd_utils" / "prompts.json"
        try:
            with open(prompts_path, 'r', encoding='utf-8') as f:
                prompts = json.load(f)
        except Exception as e:
            print(f"❌ Error loading prompts: {e}")
            return {'error': 'Could not load prompts'}
        
        # Get system prompts for mode
        system_prompts = prompts.get('system_prompts', {}).get(mode, {})
        
        agent3_system = system_prompts.get('Agent 3', '')
        agent3_user = f"""
Validate pattern compliance and clean up the content:

CONTENT TO REVIEW:
{agent1_content}

Instructions:
1. Check if content follows pattern structure
2. Limit to top 2 items if too many items listed
3. Remove excessive quotation marks
4. Verify K/M conversion with 1 decimal place
5. Return the cleaned/corrected version
"""
        
        result = self.generate_content(
            system_prompt=agent3_system,
            user_prompt=agent3_user,
            provider=provider,
            max_tokens=2500
        )
        
        if result.get('content'):
            print(f"\n✅ Agent 3 Output ({len(result['content'])} chars):")
            print("-" * 80)
            print(result['content'])
            print("-" * 80)
            print(f"✅ Tokens Used: {result['tokens']['total_tokens']}")
        
        return result
    
    def test_multi_agent(
        self,
        financial_data,
        key,
        entity_name,
        mode='english',
        provider='deepseek',
        agents='all'
    ):
        """Test multi-agent workflow for financial report generation
        
        Args:
            financial_data: Financial table data (markdown format)
            key: Financial key (e.g., 'Cash', 'AR')
            entity_name: Entity name
            mode: Language mode ('english' or 'chinese')
            provider: AI provider to use
            agents: Which agents to run ('all', '1', '2', '3', '1+2', '1+3', etc.)
            
        Returns:
            dict with agent results
        """
        print(f"\n{'='*80}")
        print(f"🧪 TESTING MULTI-AGENT WORKFLOW")
        print(f"   Key: {key}")
        print(f"   Entity: {entity_name}")
        print(f"   Mode: {mode}")
        print(f"   Provider: {provider}")
        print(f"   Agents: {agents}")
        print(f"{'='*80}\n")
        
        results = {}
        
        # Load prompts
        prompts_path = parent_dir / "fdd_utils" / "prompts.json"
        try:
            with open(prompts_path, 'r', encoding='utf-8') as f:
                prompts = json.load(f)
        except Exception as e:
            print(f"❌ Error loading prompts: {e}")
            return {'error': 'Could not load prompts'}
        
        # Get system prompts for mode
        system_prompts = prompts.get('system_prompts', {}).get(mode, {})
        
        # Determine which agents to run
        run_agent1 = agents in ['all', '1', '1+2', '1+3']
        run_agent2 = agents in ['all', '2', '1+2']
        run_agent3 = agents in ['all', '3', '1+3']
        
        # Agent 1: Content Generation
        if run_agent1:
            print("\n🤖 AGENT 1: Content Generation")
            print("-" * 80)
            
            agent1_system = system_prompts.get('Agent 1', '')
            agent1_user = f"""
Generate financial commentary for {key} - {entity_name}

FINANCIAL DATA:
{financial_data}

Instructions:
1. Analyze the financial data provided
2. Generate professional commentary following the pattern
3. Include specific figures and entities from the data
4. Keep it concise (100-120 words maximum)
5. Do NOT output raw table data
"""
            
            agent1_result = self.generate_content(
                system_prompt=agent1_system,
                user_prompt=agent1_user,
                provider=provider
            )
            
            results['agent1'] = agent1_result
            
            if agent1_result.get('content'):
                print(f"\n📝 Agent 1 Output ({len(agent1_result['content'])} chars):")
                print("-" * 80)
                print(agent1_result['content'][:500] + "..." if len(agent1_result['content']) > 500 else agent1_result['content'])
        else:
            print("\n⏭️  Skipping Agent 1")
        
        # Agent 2: Proofreading (if Agent 1 succeeded or content provided)
        if run_agent2 and results.get('agent1', {}).get('content'):
            print("\n\n🔍 AGENT 2: AI Proofreader")
            print("-" * 80)
            
            agent2_system = system_prompts.get('AI Proofreader', '')
            agent2_user = f"""
Review the following content for compliance:

ORIGINAL FINANCIAL DATA:
{financial_data}

AGENT 1 CONTENT:
{results['agent1']['content']}

Provide your review in JSON format with:
- is_compliant: boolean
- issues: array of issues found
- corrected_content: the corrected version
- figure_checks: array of figure validation results
- entity_checks: array of entity validation results
- grammar_notes: array of grammar improvements
"""
            
            agent2_result = self.generate_content(
                system_prompt=agent2_system,
                user_prompt=agent2_user,
                provider=provider,
                max_tokens=3000
            )
            
            results['agent2'] = agent2_result
            
            if agent2_result.get('content'):
                print(f"\n✅ Agent 2 Output ({len(agent2_result['content'])} chars):")
                print("-" * 80)
                print(agent2_result['content'][:500] + "..." if len(agent2_result['content']) > 500 else agent2_result['content'])
        elif run_agent2:
            print("\n⏭️  Skipping Agent 2 (no Agent 1 content)")
        else:
            print("\n⏭️  Skipping Agent 2")
        
        # Agent 3: Pattern Compliance (if Agent 1 succeeded)
        if run_agent3 and results.get('agent1', {}).get('content'):
            print("\n\n✨ AGENT 3: Pattern Compliance Validator")
            print("-" * 80)
            
            agent3_system = system_prompts.get('Agent 3', '')
            agent3_user = f"""
Validate pattern compliance and clean up the content:

CONTENT TO REVIEW:
{results['agent1']['content']}

Instructions:
1. Check if content follows pattern structure
2. Limit to top 2 items if too many items listed
3. Remove excessive quotation marks
4. Verify K/M conversion with 1 decimal place
5. Return the cleaned/corrected version
"""
            
            agent3_result = self.generate_content(
                system_prompt=agent3_system,
                user_prompt=agent3_user,
                provider=provider,
                max_tokens=2500
            )
            
            results['agent3'] = agent3_result
            
            if agent3_result.get('content'):
                print(f"\n✅ Agent 3 Output ({len(agent3_result['content'])} chars):")
                print("-" * 80)
                print(agent3_result['content'][:500] + "..." if len(agent3_result['content']) > 500 else agent3_result['content'])
        elif run_agent3:
            print("\n⏭️  Skipping Agent 3 (no Agent 1 content)")
        else:
            print("\n⏭️  Skipping Agent 3")
        
        # Summary
        print(f"\n\n{'='*80}")
        print("📊 WORKFLOW SUMMARY")
        print(f"{'='*80}")
        if run_agent1:
            print(f"Agent 1 Status: {'✅ Success' if results.get('agent1', {}).get('content') else '❌ Failed'}")
        if run_agent2:
            print(f"Agent 2 Status: {'✅ Success' if results.get('agent2', {}).get('content') else '❌ Failed/Skipped'}")
        if run_agent3:
            print(f"Agent 3 Status: {'✅ Success' if results.get('agent3', {}).get('content') else '❌ Failed/Skipped'}")
        
        total_tokens = 0
        for agent_key, agent_result in results.items():
            if agent_result.get('tokens'):
                total_tokens += agent_result['tokens']['total_tokens']
        
        print(f"Total Tokens Used: {total_tokens}")
        print(f"{'='*80}\n")
        
        return results
    
    def list_available_providers(self):
        """List all available AI providers"""
        print("\n📋 Available AI Providers:")
        print("-" * 40)
        for provider in self.clients.keys():
            print(f"  ✅ {provider}")
        print("-" * 40)
        
        return list(self.clients.keys())


if __name__ == "__main__":
    # Quick test
    print("🧪 AI Module Test")
    print("=" * 80)
    
    ai = AIModule()
    ai.list_available_providers()
    
    # Test simple generation
    result = ai.generate_content(
        system_prompt="You are a helpful assistant.",
        user_prompt="Say hello in 10 words or less.",
        provider='deepseek'
    )
    
    print(f"\n✅ Test Result:")
    print(f"Content: {result.get('content')}")
    print(f"Tokens: {result.get('tokens')}")


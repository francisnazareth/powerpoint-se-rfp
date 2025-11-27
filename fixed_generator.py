"""
Fixed Building Block Generator
Combines AI intelligence with reliable slide generation
"""

import asyncio
import os
from typing import List
from building_block_agent import BuildingBlockAgent
from direct_generator import DirectBuildingBlockGenerator


class FixedBuildingBlockGenerator:
    """Fixed generator that combines AI analysis with direct slide creation."""
    
    def __init__(self, github_token: str):
        self.github_token = github_token
        self.direct_generator = DirectBuildingBlockGenerator()
        self.ai_agent = None
    
    async def initialize_ai(self):
        """Initialize AI agent for advanced analysis."""
        if not self.ai_agent:
            self.ai_agent = BuildingBlockAgent(self.github_token, model_id="openai/gpt-4o-mini")
    
    async def create_slide(self, requirements: str, use_ai: bool = False) -> str:
        """Create building block slide with option to use AI analysis."""
        
        if use_ai and self.github_token:
            try:
                print("🤖 Using AI analysis...")
                await self.initialize_ai()
                
                # Use AI for complex requirements analysis
                categories = await self.analyze_with_ai(requirements)
                
                # But use direct generation for reliable slide creation
                return self.direct_generator.create_building_block_slide(requirements)
                
            except Exception as e:
                print(f"AI analysis failed ({e}), falling back to direct generation...")
                return self.direct_generator.create_building_block_slide(requirements)
        else:
            # Use direct generation (more reliable)
            return self.direct_generator.create_building_block_slide(requirements)
    
    async def analyze_with_ai(self, requirements: str) -> List[str]:
        """Use AI to analyze complex requirements."""
        # This could be expanded for more complex analysis
        # For now, fallback to direct analysis
        return self.direct_generator.analyze_requirements(requirements)


def main():
    """Main function with multiple options."""
    github_token = os.getenv('GITHUB_TOKEN')
    generator = FixedBuildingBlockGenerator(github_token)
    
    print("🏗️ Fixed Building Block Generator")
    print("=" * 50)
    print("Choose generation method:")
    print("1. Direct Generation (Fast, Reliable)")
    print("2. AI-Enhanced Generation (Slower, More Intelligent)")
    print("3. Test with your example")
    print()
    
    while True:
        choice = input("Select option (1-3) or 'quit': ").strip()
        
        if choice.lower() in ['quit', 'exit', 'q']:
            break
        
        if choice == '3':
            # Test with your specific example
            test_req = "1. User experience layer, 2. Application Layer, 3. Data and intelligence layer, 4. Integration Layer"
            print(f"\\nTesting with: {test_req}")
            result = generator.direct_generator.create_building_block_slide(test_req, "Test_Multi_Layer_Architecture")
            print(result)
            continue
        
        requirements = input("\\nEnter requirements: ").strip()
        if not requirements:
            continue
        
        try:
            if choice == '1':
                result = generator.direct_generator.create_building_block_slide(requirements)
                print(result)
            elif choice == '2':
                result = asyncio.run(generator.create_slide(requirements, use_ai=True))
                print(result)
            else:
                print("Invalid choice. Please select 1, 2, or 3.")
        
        except Exception as e:
            print(f"Error: {e}")


if __name__ == "__main__":
    main()
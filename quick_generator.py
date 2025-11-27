"""
Quick Building Block Generator
Simple script to generate building block slides based on requirements
"""

import asyncio
import os
import sys
from building_block_agent import BuildingBlockAgent


async def generate_slide(requirements: str):
    """Generate a building block slide for given requirements."""
    github_token = os.getenv('GITHUB_TOKEN')
    if not github_token:
        print("❌ GITHUB_TOKEN environment variable not set")
        return
    
    agent = BuildingBlockAgent(github_token, model_id="openai/gpt-4o-mini")
    
    print(f"🏗️ Generating building block slide...")
    print(f"Requirements: {requirements}")
    print("=" * 60)
    
    result = await agent.create_building_block_presentation(requirements)
    print("\n✅ Building block slide generated!")
    return result


def main():
    """Main function to handle command line arguments or interactive input."""
    if len(sys.argv) > 1:
        # Use command line argument
        requirements = " ".join(sys.argv[1:])
        asyncio.run(generate_slide(requirements))
    else:
        # Interactive mode
        print("🏗️ Quick Building Block Generator")
        print("=" * 40)
        print("Enter your requirements below:")
        print()
        
        requirements = input("Requirements: ").strip()
        
        if requirements:
            asyncio.run(generate_slide(requirements))
        else:
            print("No requirements provided.")


if __name__ == "__main__":
    main()
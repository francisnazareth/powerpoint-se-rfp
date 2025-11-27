"""
Example usage of the PowerPoint AI Agent
Demonstrates different ways to interact with the agent programmatically
"""

import asyncio
import os
from powerpoint_agent import PowerPointAgent


async def example_web_app_architecture():
    """Example 1: Create a web application architecture presentation"""
    print("🌐 Example 1: Web Application Architecture")
    print("=" * 50)
    
    # Get GitHub token from environment
    github_token = os.getenv('GITHUB_TOKEN')
    if not github_token:
        print("❌ Please set GITHUB_TOKEN environment variable")
        return
    
    # Initialize agent
    agent = PowerPointAgent(github_token, model_id="openai/gpt-4.1")
    
    # Create presentation for web app architecture
    request = """Create a comprehensive presentation about building a scalable web application on Azure.
    
    Include the following:
    1. Title slide with the presentation topic
    2. Overview of the web application architecture
    3. Frontend hosting with Azure App Service
    4. Backend API with Azure Functions
    5. Database layer with Azure SQL Database
    6. Security with Azure Key Vault
    7. Monitoring with Azure Application Insights
    8. Architecture diagram showing all components
    
    Make it professional and suitable for a technical audience."""
    
    await agent.create_presentation(request)
    print("\n✅ Web app architecture presentation completed!\n")


async def example_ai_chatbot_solution():
    """Example 2: Create an AI chatbot solution presentation"""
    print("🤖 Example 2: AI Chatbot Solution")
    print("=" * 50)
    
    github_token = os.getenv('GITHUB_TOKEN')
    if not github_token:
        print("❌ Please set GITHUB_TOKEN environment variable")
        return
    
    agent = PowerPointAgent(github_token, model_id="openai/gpt-4.1-mini")  # Using mini for faster response
    
    request = """Design a presentation for an intelligent chatbot solution on Azure.
    
    Cover these aspects:
    1. Title: "AI-Powered Customer Service Chatbot"
    2. Business case and benefits
    3. Azure OpenAI Service for natural language processing
    4. Azure Cosmos DB for conversation history and user data
    5. Azure Bot Service for chatbot framework
    6. Azure Application Gateway for load balancing
    7. Azure Cognitive Services for additional AI capabilities
    8. Architecture diagram of the complete solution
    9. Implementation roadmap and next steps
    
    Focus on business value and technical implementation."""
    
    await agent.create_presentation(request)
    print("\n✅ AI chatbot solution presentation completed!\n")


async def example_data_analytics_pipeline():
    """Example 3: Create a data analytics pipeline presentation"""
    print("📊 Example 3: Data Analytics Pipeline")
    print("=" * 50)
    
    github_token = os.getenv('GITHUB_TOKEN')
    if not github_token:
        print("❌ Please set GITHUB_TOKEN environment variable")
        return
    
    agent = PowerPointAgent(github_token)
    
    request = """Create a presentation about building a modern data analytics pipeline on Azure.
    
    Structure the presentation as follows:
    1. Title slide: "Modern Data Analytics on Azure"
    2. Data ingestion challenges and solutions
    3. Azure Data Factory for ETL processes
    4. Azure Storage Account for data lake storage
    5. Azure Synapse Analytics for data warehousing
    6. Azure Analysis Services for data modeling
    7. Power BI for visualization and reporting
    8. Azure Machine Learning for predictive analytics
    9. Complete architecture diagram
    10. Benefits and ROI considerations
    
    Make it suitable for both technical and business stakeholders."""
    
    await agent.create_presentation(request)
    print("\n✅ Data analytics pipeline presentation completed!\n")


async def example_microservices_architecture():
    """Example 4: Create a microservices architecture presentation"""
    print("🔧 Example 4: Microservices Architecture")
    print("=" * 50)
    
    github_token = os.getenv('GITHUB_TOKEN')
    if not github_token:
        print("❌ Please set GITHUB_TOKEN environment variable")
        return
    
    agent = PowerPointAgent(github_token)
    
    request = """Develop a presentation about implementing microservices architecture on Azure.
    
    Include these key topics:
    1. Title: "Microservices Architecture on Azure"
    2. Introduction to microservices principles
    3. Azure Kubernetes Service (AKS) for container orchestration
    4. Azure API Management for service gateway
    5. Azure Service Bus for inter-service communication
    6. Azure Container Registry for image management
    7. Azure DevOps for CI/CD pipelines
    8. Azure Monitor for distributed tracing
    9. Architecture diagram showing service interactions
    10. Migration strategy from monolithic applications
    
    Target audience: Software architects and development teams."""
    
    await agent.create_presentation(request)
    print("\n✅ Microservices architecture presentation completed!\n")


async def run_all_examples():
    """Run all examples sequentially"""
    print("🚀 Running All PowerPoint Agent Examples")
    print("=" * 60)
    print("This will create 4 different presentations demonstrating")
    print("various Azure architecture patterns and solutions.\n")
    
    examples = [
        example_web_app_architecture,
        example_ai_chatbot_solution,
        example_data_analytics_pipeline,
        example_microservices_architecture
    ]
    
    for i, example in enumerate(examples, 1):
        try:
            print(f"Running example {i} of {len(examples)}...")
            await example()
            print("Waiting 2 seconds before next example...\n")
            await asyncio.sleep(2)  # Brief pause between examples
        except Exception as e:
            print(f"❌ Error in example {i}: {str(e)}\n")
            continue
    
    print("🎉 All examples completed!")
    print("Check your current directory for the generated .pptx files.")


async def interactive_example():
    """Interactive example allowing user input"""
    print("💬 Interactive PowerPoint Agent Example")
    print("=" * 50)
    
    github_token = os.getenv('GITHUB_TOKEN')
    if not github_token:
        print("❌ Please set GITHUB_TOKEN environment variable")
        return
    
    agent = PowerPointAgent(github_token)
    
    print("Enter your custom request for a PowerPoint presentation.")
    print("Examples:")
    print("- 'Create slides about IoT architecture with IoT Hub and Stream Analytics'")
    print("- 'Build a presentation on Azure security best practices'")
    print("- 'Design slides for a machine learning pipeline on Azure'\n")
    
    user_input = input("Your request: ").strip()
    
    if user_input:
        print(f"\n🔄 Creating presentation based on your request...\n")
        await agent.create_presentation(user_input)
        print("\n✅ Your custom presentation is ready!")
    else:
        print("No request provided.")


def main_menu():
    """Display main menu and handle user selection"""
    while True:
        print("\n" + "=" * 60)
        print("🎯 PowerPoint AI Agent - Example Menu")
        print("=" * 60)
        print("1. Web Application Architecture")
        print("2. AI Chatbot Solution")
        print("3. Data Analytics Pipeline")
        print("4. Microservices Architecture")
        print("5. Run All Examples")
        print("6. Interactive Custom Request")
        print("7. Exit")
        print("=" * 60)
        
        choice = input("Select an option (1-7): ").strip()
        
        if choice == '1':
            asyncio.run(example_web_app_architecture())
        elif choice == '2':
            asyncio.run(example_ai_chatbot_solution())
        elif choice == '3':
            asyncio.run(example_data_analytics_pipeline())
        elif choice == '4':
            asyncio.run(example_microservices_architecture())
        elif choice == '5':
            asyncio.run(run_all_examples())
        elif choice == '6':
            asyncio.run(interactive_example())
        elif choice == '7':
            print("👋 Goodbye!")
            break
        else:
            print("❌ Invalid choice. Please select 1-7.")


if __name__ == "__main__":
    # Check for GitHub token
    if not os.getenv('GITHUB_TOKEN'):
        print("⚠️  GitHub Token Setup Required")
        print("=" * 50)
        print("Please set your GitHub Personal Access Token as an environment variable:")
        print("PowerShell: $env:GITHUB_TOKEN='your_token_here'")
        print("Bash: export GITHUB_TOKEN='your_token_here'")
        print("\nCreate a token at: https://github.com/settings/tokens")
        print("No additional permissions are required beyond the default.")
        exit(1)
    
    main_menu()
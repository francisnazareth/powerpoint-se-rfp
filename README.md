# PowerPoint AI Agent

A Python AI agent that creates PowerPoint presentations using Azure building blocks. Built with Microsoft Agent Framework and GitHub models.

## Features

- 🤖 **AI-Powered**: Uses GitHub models (GPT-4.1) for intelligent content generation
- 🏗️ **Azure-Focused**: Comprehensive knowledge of Azure services and architecture patterns
- 📊 **PowerPoint Generation**: Creates professional PPTX files with multiple slide types
- 🎨 **Architecture Diagrams**: Generates visual diagrams of Azure components
- 🔧 **Tool Integration**: Leverages Microsoft Agent Framework for robust agent capabilities

## Azure Building Blocks Supported

The agent has extensive knowledge of Azure services across these categories:

- **Compute**: Virtual Machines, App Service, Azure Functions, Container Instances, AKS
- **Storage**: Blob Storage, Disk Storage, Files, Queue Storage, Table Storage
- **Networking**: Virtual Network, Load Balancer, Application Gateway, VPN Gateway, ExpressRoute
- **Database**: SQL Database, Cosmos DB, PostgreSQL, MySQL, Redis Cache
- **AI/ML**: Cognitive Services, Machine Learning, AI Search, OpenAI Service, Bot Service
- **Security**: Key Vault, Active Directory, Security Center, Sentinel, Firewall
- **Monitoring**: Azure Monitor, Log Analytics, Application Insights, Service Health, Advisor

## Prerequisites

1. **Python 3.8+**
2. **GitHub Personal Access Token** with access to GitHub Models
   - Create one at: https://github.com/settings/tokens
   - Set as environment variable: `GITHUB_TOKEN`

## Installation

1. Install required packages:
   ```powershell
   pip install agent-framework-azure-ai --pre
   pip install python-pptx
   ```

   > **Important**: The `--pre` flag is required while Agent Framework is in preview.

2. Set your GitHub token as an environment variable:
   ```powershell
   $env:GITHUB_TOKEN="your_github_token_here"
   ```

## Usage

### Basic Usage

```python
import asyncio
from powerpoint_agent import PowerPointAgent

async def create_presentation():
    # Initialize the agent
    agent = PowerPointAgent(github_token="your_token")
    
    # Create a presentation
    request = "Create a presentation about a web application architecture on Azure with App Service, SQL Database, and Key Vault"
    await agent.create_presentation(request)

asyncio.run(create_presentation())
```

### Interactive Mode

Run the script directly for interactive mode:

```powershell
python powerpoint_agent.py
```

## Example Requests

Here are some example requests you can make to the agent:

1. **Web Application Architecture**:
   ```
   Create a presentation about a web application architecture on Azure with App Service, SQL Database, and Key Vault
   ```

2. **AI Chatbot Solution**:
   ```
   Build slides for an AI-powered chatbot solution using Azure OpenAI, Cosmos DB, and Application Gateway
   ```

3. **Data Analytics Pipeline**:
   ```
   Generate a presentation about Azure data analytics pipeline with Storage Account, Data Factory, and Synapse Analytics
   ```

4. **Microservices Architecture**:
   ```
   Create slides showing a microservices architecture with AKS, API Management, and Service Bus
   ```

## Agent Capabilities

The agent can perform the following actions:

- **`get_azure_services`**: Retrieve information about Azure services by category
- **`create_powerpoint_slide`**: Create individual slides with title and content
- **`create_architecture_diagram`**: Generate visual architecture diagrams with Azure components
- **`save_presentation`**: Save the completed presentation as a PPTX file

## Output

The agent generates:
- Professional PowerPoint presentations (.pptx files)
- Title slides with branding
- Content slides with detailed explanations
- Architecture diagrams with colored component boxes
- Properly formatted text and layouts

## Architecture

```
PowerPointAgent
├── Microsoft Agent Framework (Core)
├── OpenAI Client (GitHub Models)
├── Azure Building Blocks Knowledge Base
├── PowerPoint Generation Tools
│   ├── Slide Creation
│   ├── Architecture Diagrams
│   └── Presentation Management
└── Interactive Interface
```

## Models Used

- **Default**: `openai/gpt-4.1` (GitHub Models)
- **Alternative**: Any supported GitHub model can be configured

## Configuration

You can customize the agent by modifying:

- **Model selection**: Change `model_id` parameter
- **Azure services**: Extend `azure_building_blocks` dictionary
- **Slide templates**: Modify PowerPoint layouts
- **Colors and styling**: Update visual formatting

## Error Handling

The agent includes robust error handling for:
- Invalid GitHub tokens
- Network connectivity issues
- PowerPoint generation errors
- File system permissions

## Best Practices

1. **Use descriptive requests**: Be specific about the Azure services and architecture you want
2. **Include context**: Mention the use case or business scenario
3. **Review generated content**: The agent creates drafts that you can refine
4. **Save frequently**: Use the save function to preserve your work

## Limitations

- Maximum 6 components per architecture diagram (for visual clarity)
- Requires active internet connection for AI model access
- GitHub Models rate limits apply (free tier available)
- PowerPoint features are limited to basic shapes and text

## Contributing

This is an example implementation. You can extend it by:
- Adding more Azure services to the knowledge base
- Implementing advanced PowerPoint features
- Adding custom slide templates
- Integrating with other AI models
- Adding validation and error recovery

## License

This project is for educational and demonstration purposes. Please ensure compliance with GitHub Models terms of service and OpenAI usage policies.
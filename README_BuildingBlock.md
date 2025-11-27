# Building Block PowerPoint Agent

A specialized Python AI agent that creates single-slide PowerPoint presentations following the **ConceptualArchitecture-Samples.pptx** building block format. Built with Microsoft Agent Framework and designed to follow your specific Azure service recommendations.

## 🎯 Key Features

- **Single-slide focus**: Creates one comprehensive building block slide per requirement set
- **Format compliance**: Follows your ConceptualArchitecture-Samples.pptx structure
- **Smart service recommendations**: Automatically suggests appropriate Azure services based on requirements
- **Category-based organization**: Organizes services into logical building blocks with color coding
- **AI-powered analysis**: Uses GPT models to understand requirements and recommend optimal architectures

## 🏗️ Service Recommendations by Category

### AI & Analytics Requirements
- **Primary Services**: Azure OpenAI, Microsoft Fabric, Azure Databricks
- **Supporting Services**: Azure AI Services, Azure Synapse Analytics, Azure ML, Power BI, Azure AI Search

### Web Application Requirements
- **Primary Services**: Azure Web Apps, Azure Container Apps, Azure Kubernetes Service
- **Supporting Services**: Azure App Service, Azure Front Door, Azure Application Gateway, Azure Load Balancer

### Data Platform Requirements
- **Primary Services**: Azure SQL Database, Azure Cosmos DB, Azure Storage
- **Supporting Services**: Azure Data Factory, Azure Data Lake, Azure Synapse, Azure Purview

### Integration Requirements
- **Primary Services**: Azure API Management, Azure Service Bus, Azure Logic Apps
- **Supporting Services**: Azure Event Grid, Azure Event Hub, Function Apps, Power Automate

### Security Requirements
- **Primary Services**: Microsoft Entra ID, Azure Key Vault, Microsoft Sentinel
- **Supporting Services**: Azure Firewall, Microsoft Defender, Azure Policy, Azure Monitor

### Infrastructure Requirements
- **Primary Services**: Azure Virtual Networks, Azure Virtual Machines, Azure Backup
- **Supporting Services**: Azure DevOps, Azure Monitor, Azure Policy and Compliance

## 🚀 Quick Start

### Prerequisites
1. **Python 3.8+**
2. **GitHub Personal Access Token** (already set in your environment)
3. **Required packages** (already installed):
   ```bash
   pip install agent-framework-azure-ai --pre
   pip install python-pptx
   ```

### Usage Methods

#### Method 1: Interactive Mode
```powershell
python building_block_agent.py
```
Follow the prompts to enter your requirements.

#### Method 2: Quick Generator
```powershell
python quick_generator.py "AI-powered analytics platform with web interface"
```

#### Method 3: Command Line
```powershell
python quick_generator.py AI analytics web application database integration
```

## 📋 Example Requirements

### AI & Analytics
```
"AI-powered customer service solution with chat capabilities and analytics"
"Machine learning platform for predictive analytics and data processing"
"Intelligent document processing system with OCR and NLP"
```

### Web Applications
```
"Modern web application with microservices architecture and database"
"E-commerce platform with payment processing and inventory management"
"Multi-tenant SaaS application with user management"
```

### Complex Solutions
```
"Enterprise data platform with AI analytics, web portal, and real-time processing"
"Government e-invoicing solution with integration, security, and compliance"
"Healthcare management system with patient portal, analytics, and mobile access"
```

## 🎨 Building Block Structure

The agent creates slides with:

1. **Title Section**: Clear slide title at the top
2. **Building Block Grid**: 3-column layout with colored blocks representing:
   - **Purple**: AI & Analytics services
   - **Blue**: Web Application services  
   - **Teal**: Data Platform services
   - **Orange**: Integration services
   - **Red**: Security services
   - **Dark Blue**: Infrastructure services
3. **Service Details**: Each block shows category name and top 3 primary services
4. **Requirements Summary**: Original requirements displayed at bottom

## 📁 Generated Files

Files are saved with descriptive names like:
- `Solution_Architecture_Building_Blocks.pptx`
- `AI_Powered_Analytics_Platform_Architecture.pptx`
- `Solution_Architecture_Building_Blocks_Data_and_Intelligence_Layer.pptx`

## 🔧 Customization

### Adding New Service Categories
Edit the `service_recommendations` dictionary in `BuildingBlockAgent` class:

```python
self.service_recommendations["new_category"] = {
    "primary": ["Service 1", "Service 2", "Service 3"],
    "supporting": ["Service 4", "Service 5", "Service 6"]
}
```

### Changing Colors
Modify the `building_block_colors` dictionary:

```python
self.building_block_colors["new_category"] = RGBColor(R, G, B)
```

### Adjusting Layout
Modify block positioning in `create_building_block_slide` method:
- `blocks_per_row`: Number of blocks per row
- `block_width`/`block_height`: Block dimensions  
- `spacing_x`/`spacing_y`: Space between blocks

## 🤖 AI Agent Workflow

1. **Requirements Analysis**: Agent analyzes input text to identify solution categories
2. **Service Recommendation**: Matches categories to predefined Azure service sets
3. **Slide Creation**: Generates building block layout with appropriate services
4. **File Generation**: Saves complete PowerPoint presentation

## 📊 Comparison with Original Agent

| Feature | Original PowerPoint Agent | Building Block Agent |
|---------|--------------------------|---------------------|
| **Slides per request** | Multiple (5-10) | Single comprehensive |
| **Format** | Standard PowerPoint layouts | Building block format |
| **Service selection** | AI-generated recommendations | Predefined category-based |
| **Visual structure** | Text-heavy slides | Visual building blocks |
| **Use case** | Detailed presentations | Architecture overview |

## 🎯 Best Practices

1. **Be specific with requirements**: Include technology preferences and use cases
2. **Mention multiple domains**: For complex solutions, specify all needed areas
3. **Review generated blocks**: Verify the selected services match your needs
4. **Combine requirements**: Single request for comprehensive solutions works better
5. **Use descriptive language**: Natural language descriptions work best

## 💡 Tips for Requirements

### Good Examples:
✅ "AI-powered e-commerce platform with real-time analytics and mobile app"  
✅ "Healthcare management system with patient portal, data analytics, and integration to external systems"  
✅ "Financial services platform with fraud detection, web portal, and regulatory compliance"

### Less Effective:
❌ "Web app"  
❌ "Database system"  
❌ "AI solution"

## 🔍 Troubleshooting

### Common Issues:
1. **No slides generated**: Check GITHUB_TOKEN environment variable
2. **Missing services**: Ensure requirements mention relevant keywords  
3. **Layout issues**: Verify PowerPoint file opens correctly in PowerPoint application
4. **File not found**: Check current working directory for generated .pptx files

### Debug Mode:
Add print statements to track agent decisions:
```python
print(f"Identified categories: {needed_categories}")
print(f"Selected services: {services}")
```

## 🎉 Success Metrics

The agent successfully generates building block slides when:
- ✅ Requirements clearly identify 2-4 solution categories
- ✅ Generated services align with your predefined recommendations
- ✅ Slide layout follows the ConceptualArchitecture-Samples.pptx format
- ✅ All building blocks are properly color-coded and labeled
- ✅ PowerPoint file opens correctly with professional formatting

---

This building block agent provides a focused, format-compliant solution for creating architecture overview slides that follow your established building block methodology while leveraging AI intelligence for service recommendations.
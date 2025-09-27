# Excel Sample Data Generator

A Streamlit-based web application that generates realistic employee data and exports it to Excel, complete with AI-powered insights and professional PDF reporting. This tool is perfect for testing, demonstrations, and training purposes where realistic but synthetic employee data is needed.

## ✨ Features

### 📊 Data Generation
- Generate realistic employee records with comprehensive details:
  - Personal information (name, email, phone, address)
  - Employment details (ID, department, position, manager, hire date)
  - Compensation data (salary, bonus, stock options)
  - Performance metrics and ratings
  - Team assignments and project allocations

### 🛠 Customization Options
- **Data Volume Control**:
  - Generate any number of records (default: 100)
  - Control data distribution across departments and positions
  - Adjust duplicate record frequency (0-100%)
  - Configure missing data percentage for realistic testing

- **ID Generation**:
  - Custom ID prefixes (e.g., 'EMP', 'DEV')
  - Configurable ID ranges and formats
  - Weighted random prefix selection
  - Duplicate ID generation control

### 🤖 AI-Powered Analysis
- Comprehensive workforce analytics:
  - Department and position distribution
  - Compensation analysis by role and seniority
  - Performance metrics and trends
  - Team composition and structure
  - Data quality assessment

### 📑 Professional Reporting
- **Excel Export**:
  - Multiple formatted worksheets
  - Data validation and formatting
  - Summary statistics and visualizations
  - Clean, professional layout

- **PDF Reports**:
  - Executive summary
  - Detailed analysis sections
  - Visual charts and graphs
  - Professional typography and styling

### 🔄 Advanced Features
- Team and project management
- Performance tracking
- Compensation analysis
- Departmental insights
- Data validation rules

## 🚀 Recent Updates

### Added
- Comprehensive team and project management features
- Advanced performance tracking metrics
- Detailed compensation analysis by role and department
- Enhanced data validation and error handling
- Support for multiple Excel sheets in a single workbook
- Professional PDF reporting with visualizations
- Performance optimizations for large datasets

## 🛠️ Prerequisites

- Python 3.8+
- Ollama (for AI features)
- Required Python packages (install via `pip install -r requirements.txt`):
  - streamlit
  - pandas
  - numpy
  - faker
  - xlsxwriter
  - reportlab
  - matplotlib
  - seaborn
  - ollama
  - faker
  - openpyxl
  - reportlab
  - python-dotenv
  - ollama
  - matplotlib
  - seaborn
  - xlsxwriter

## Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/yourusername/your-repo.git
   cd your-repo/Excel\ Sample\ Data\ Generator
   ```

2. Install the required packages:
   ```bash
   pip install -r requirements.txt
   ```

3. Set up Ollama (for AI insights):
   - Install Ollama from [ollama.ai](https://ollama.ai/)
   - Pull the required model: `ollama pull gemma3:4b`

## Usage

1. Run the Streamlit app:
   ```bash
   streamlit run generate_employee_data.py
   ```

2. In the web interface:
   - Adjust the number of employees using the slider
  - Click "Generate Employee Data" to create the dataset
  - Use the "Export to Excel" button to download the data
  - Generate a PDF report with the "Generate PDF Report" button

## Performance Considerations

### What Can Make It Slow

1. **Large Datasets**:
   - Generating more than 10,000 records may cause performance issues
   - PDF generation is particularly resource-intensive with large datasets

2. **AI Insights**:
   - The AI analysis can take 30-60 seconds depending on your hardware
   - Internet connection quality affects response time from the Ollama server

3. **System Resources**:
   - Running out of memory with very large datasets
   - Multiple concurrent users may slow down the application

### Potential Crash Scenarios

1. **Missing Dependencies**:
   - Ensure all required Python packages are installed
   - Verify Ollama is running if using AI features

2. **File Permissions**:
   - The app needs write permissions in the current directory
   - Check for sufficient disk space before generating large files

3. **Memory Issues**:
   - Generating very large Excel files (>100MB) may cause memory errors
   - PDF generation with many pages can be memory-intensive

## Best Practices

1. **For Large Datasets**:
   - Generate data in smaller batches (e.g., 10,000 records at a time)
   - Use the command-line version for headless operation
   - Close other memory-intensive applications
   - Consider increasing Python's memory limits if needed

2. **For Better Performance**:
   - Use the latest version of all dependencies
   - Run the application on a machine with sufficient RAM (8GB+ recommended for large datasets)
   - Use a virtual environment to avoid dependency conflicts
   - Disable AI analysis when not needed for faster generation

3. **For Realistic Data**:
   - Adjust the duplicate percentage based on your testing needs
   - Use the missing data feature to test data quality scenarios
   - Customize department and role distributions to match your organization

4. **Troubleshooting**:
   - Check the console for error messages
   - Verify that Ollama is running if AI features aren't working
   - Clear the browser cache if the UI becomes unresponsive
   - Check disk space before generating large files
   - Ensure you have write permissions in the output directory

## About This Project

> "I wanted to make testing easier, so I created a script that generates realistic test subject data with random information. This tool was born out of the need for quick, realistic datasets for development and testing purposes."

## Version History

### v1.0.0 (2024-09-23)
- Initial release with basic employee data generation
- Excel and PDF export capabilities
- Basic AI-powered analysis

### v1.1.0 (2025-09-21)
- Added support for multiple Excel sheets
- Enhanced data generation with more realistic distributions
- Improved duplicate record handling
- Better error handling and validation
- Performance optimizations
- More detailed reporting and analysis

## 📜 Version History

### v2.3.0 (Latest)
- Added comprehensive team and project management features
- Implemented advanced performance tracking metrics
- Enhanced compensation analysis with role and department breakdowns
- Improved PDF report generation with professional styling
- Added support for multiple output formats (Excel, PDF with Summary)
- Optimized performance for large datasets

### v2.2.0
- Added AI-powered analysis using Ollama
- Implemented detailed data validation rules
- Enhanced error handling and user feedback
- Added support for custom ID generation with prefixes
- Improved duplicate record handling

### v2.1.0
- Added support for multiple Excel sheets in a single workbook
- Implemented realistic salary distribution based on roles
- Enhanced data generation with more realistic employee attributes
- Added support for team assignments and project allocations
- Improved handling of missing data scenarios

### v2.0.0
- Complete rewrite with Streamlit for web interface
- Added professional PDF reporting with ReportLab
- Implemented data visualization with Matplotlib/Seaborn
- Added comprehensive data analysis features
- Improved data generation algorithms

### v1.0.0
- Initial release
- Basic employee data generation
- Excel export functionality
- Simple data customization options

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request. Here are some ways you can contribute:
- Report bugs
- Suggest new features
- Improve documentation
- Submit code improvements

---

*Last Updated: September 27, 2024*


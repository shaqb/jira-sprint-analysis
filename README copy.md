# Jira Sprint Analysis Tool

A Python-based tool for analyzing Jira sprint data, comparing original estimates against AI-generated estimates and actual time spent across different disciplines (BA, FE, TA, QA).

## Features

- Interactive web interface using Streamlit
- Processes Jira export data to analyze estimation accuracy
- Compares original estimates, AI-generated estimates, and actual time spent
- Breaks down analysis by discipline (BA, FE, TA, QA)
- Generates detailed Excel report with:
  - Summary Dashboard
  - Analysis Results
  - Data Visualizations
- Creates interactive visualizations for estimate comparisons
- Handles missing data gracefully with detailed reporting

## Quick Start

1. Clone the repository:
```bash
git clone <repository-url>
cd halfords-jira-sprint-analysis
```

2. Create and activate a virtual environment:
```bash
python -m venv venv
source venv/bin/activate  # On Windows, use: venv\Scripts\activate
```

3. Install dependencies:
```bash
pip install -r requirements.txt
```

4. Run the Streamlit app:
```bash
streamlit run app.py
```

5. Upload your Jira export Excel file through the web interface or place it in the `input` directory.

## Project Structure

```
halfords-jira-sprint-analysis/
├── jira_analysis/           # Main package directory
│   ├── __init__.py         # Package initialization
│   ├── analyzer.py         # Core analysis functionality
│   ├── visualization.py    # Plotting and Excel formatting
│   └── utils.py           # Utility functions
├── tests/                  # Test directory
│   ├── __init__.py
│   ├── test_analyzer.py
│   ├── test_visualization.py
│   └── test_utils.py
├── input/                  # Place Jira export files here
├── output/                # Generated reports appear here
├── app.py                 # Streamlit web application
├── requirements.txt       # Production dependencies
├── requirements-dev.txt   # Development dependencies
├── setup.py              # Package setup file
├── pyproject.toml        # Build system requirements
├── tox.ini              # Tox configuration
└── README.md
```

## Development

1. Install development dependencies:
```bash
pip install -r requirements-dev.txt
```

2. Install the package in development mode:
```bash
pip install -e .
```

3. Run tests:
```bash
python -m pytest
```

## Input File Format

The tool expects a Jira export Excel file with the following columns:
- Key: Jira ticket key
- For each discipline (QA, TA, FE, BA):
  - Original Estimate based on old way
  - Revised Estimate based on AI
  - Actual Time based on AI

## Output

The tool generates:
1. Interactive web interface with:
   - Data preview
   - Analysis results
   - Interactive visualizations
2. Excel report containing:
   - Summary statistics
   - Detailed analysis by discipline
   - Visualizations
   - Data quality metrics

## Contributing

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add some amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## License

This project is licensed under the MIT License - see the LICENSE file for details.

# Medical Case Analysis with AI

This project uses AI to analyze medical cases and provide treatment recommendations based on clinical data and coronary CTA reports.

## Features

- Processes medical cases from Excel files
- Uses AI to analyze coronary CTA reports
- Provides treatment recommendations based on latest medical guidelines
- Supports concurrent processing
- Includes error handling and retry mechanisms
- Generates detailed analysis reports

## Requirements

- Python 3.8+
- OpenAI API access
- Required Python packages (see `requirements.txt`)

## Installation

1. Clone the repository:
```bash
git clone https://github.com/yourusername/medical-case-analysis.git
cd medical-case-analysis
```

2. Install required packages:
```bash
pip install -r requirements.txt
```

3. Configure your OpenAI API credentials in the code:
```python
def create_client():
    return OpenAI(
        base_url="YOUR_BASE_URL",
        api_key="YOUR_API_KEY",
        timeout=60
    )
```

## Usage

1. Prepare your input Excel file with the following columns:
   - AGE
   - SEX
   - chief complaint
   - present history
   - past history
   - DAY
   - CTA

2. Run the analysis:
```bash
python medical_analysis.py
```

3. Check the results in:
   - `medical_analysis_results.xlsx`
   - `medical_analysis.log`

## Output Format

The program generates an Excel file with the following columns:
- index
- gender
- age
- surgery_date
- model_output
- status
- attempt_count

## Error Handling

- Automatically retries on API failures
- Creates CSV backup if Excel save fails
- Detailed logging for troubleshooting

## Contributing

Feel free to open issues or submit pull requests for improvements.

## License

MIT License - see LICENSE file for details 
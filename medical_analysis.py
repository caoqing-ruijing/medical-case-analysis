import pandas as pd
from openai import OpenAI, APIError, APIConnectionError
from tqdm import tqdm
import concurrent.futures
from tenacity import retry, stop_after_attempt, wait_exponential, retry_if_exception_type
import time
import logging
import os 

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("medical_analysis.log"),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

def create_client():
    """Create OpenAI client with configuration"""
    return OpenAI(
        base_url="",
        api_key="",
        timeout=60
    )

PROMPT_TEMPLATE = '''
You are a cardiac intervention specialist. Based on the patient's clinical data and coronary CTA report, 
please determine whether PCI (Percutaneous Coronary Intervention) is needed according to the latest 
cardiovascular intervention guidelines before the patient's surgery date, and recommend the best treatment plan.
Please follow the process strictly:

**Patient Information**
Age: <age>{{AGE}}</age>
Gender: <gender>{{GENDER}}</gender>
Chief Complaint: <chief complaint>{{CHIEF COMPLAINT}}</chief complaint>
Present History: <present history>{{PRESENT HISTORY}}</present history>
Past History: <past history>{{PAST HISTORY}}</past history>
Surgery Date: <day>{{DAY}}</day>
Coronary CTA Report:
<CTA>
{{CORONARY_CTA}}
</CTA>

**Analysis Process**

1. Anatomical Feature Analysis (Required):
   - List in <analysis>:
     * Lesion Location (Left Main/LAD/LCX/RCA etc.)
     * Stenosis Degree (Percentage)
     * Lesion Type (A/B1/B2/C)
     * Special Features (Calcification/Thrombus/Bifurcation etc.)

2. Indication Assessment (Must cite guidelines):
   - Reference latest cardiovascular intervention guidelines before surgery date
     * Cite specific provisions (Format: "ESC NSTE-ACS Guidelines Chapter X Item Y")
   - Assess if meeting following indications:
     * Acute Coronary Syndrome
     * High-risk Chronic Coronary Syndrome
     * Significant Ischemic Evidence
     * Left Main Disease ≥50%
     * Proximal LAD Stenosis ≥70% etc.

3. Treatment Decision:
   - If PCI needed, select based on lesion characteristics (with reasons):
     * Balloon Dilation
     * Drug-Eluting Stent Implantation
     * Cutting Balloon
     * Drug-Coated Balloon

**Required Output Format**
<decision>
Treatment Decision: [1/0]
</decision>
<basis>
[Brief explanation based on specific guideline provisions]
</basis>
<recommendation>
Recommended Plan: [Specific procedure]
Based on: [Guideline name] Chapter [X]
Reason: [Combined with lesion characteristics]
</recommendation>
'''

@retry(stop=stop_after_attempt(5),
       wait=wait_exponential(multiplier=1, min=2, max=60),
       retry=retry_if_exception_type((APIError, APIConnectionError)))
def api_call_with_retry(client, prompt):
    """Make API call with retry mechanism"""
    return client.chat.completions.create(
        model="o3-mini",
        messages=[{"role": "user", "content": prompt}],
        max_tokens=5000
    )

def process_case(args):
    """Process a single medical case"""
    index, row = args
    client = create_client()
    error_logs = []

    processed_prompt = PROMPT_TEMPLATE \
        .replace("{{AGE}}", str(row.get('AGE', 'Unknown'))) \
        .replace("{{GENDER}}", str(row.get('SEX', 'Unknown'))) \
        .replace("{{CHIEF COMPLAINT}}", str(row.get('chief complaint', 'Unknown'))) \
        .replace("{{PRESENT HISTORY}}", str(row.get('present history', 'Unknown'))) \
        .replace("{{PAST HISTORY}}", str(row.get('past history', 'Unknown'))) \
        .replace("{{DAY}}", str(row.get('DAY', 'Unknown'))) \
        .replace("{{CORONARY_CTA}}", str(row.get('CTA', 'Unknown'))[:2000])

    for attempt in range(3):
        try:
            response = api_call_with_retry(client, processed_prompt)
            raw_content = response.choices[0].message.content
            if raw_content.strip():
                cleaned_content = raw_content.replace("</decision>", "</decision>") \
                    .replace("</recommendation>", "</recommendation>")

                format_warning = ""
                if not all(tag in cleaned_content for tag in ["<decision>", "<recommendation>"]):
                    format_warning = "(Format Warning: Missing Required Tags)"
                elif not all(tag in cleaned_content for tag in ["</decision>", "</recommendation>"]):
                    format_warning = "(Format Warning: Unclosed Tags)"

                return {
                    "index": index,
                    "gender": row['SEX'],
                    "age": row['AGE'],
                    "surgery_date": row['DAY'],
                    "model_output": f"{cleaned_content}{format_warning}",
                    "status": "success",
                    "attempt_count": attempt + 1
                }
            raise ValueError("Empty Response")

        except Exception as e:
            error_logs.append(f"Attempt {attempt + 1}: {type(e).__name__} - {str(e)}")
            if "rate limit" in str(e).lower():
                time.sleep(15)
            continue

    logger.error(f"Case {index} failed: {'; '.join(error_logs)}")
    return {
        "index": index,
        "status": "failed",
        "error_log": error_logs[-1] if error_logs else "Unknown Error"
    }

def main():
    """Main execution function"""
    config = {
        "input_path": "END.xlsx",
        "output_path": "medical_analysis_results.xlsx",
        "max_workers": 1,
        "rate_limit_delay": 150
    }

    df = pd.read_excel(config["input_path"]).convert_dtypes()
    logger.info(f"Successfully loaded {len(df)} cases")

    with concurrent.futures.ThreadPoolExecutor(
            max_workers=config["max_workers"]
    ) as executor:
        futures = []
        results = []

        for idx, row in df.iterrows():
            future = executor.submit(process_case, (idx, row))
            futures.append(future)

        with tqdm(total=len(futures), desc="Medical Analysis Progress") as progress:
            for future in concurrent.futures.as_completed(futures):
                results.append(future.result())
                progress.update(1)
                if len(results) % 20 == 0:
                    time.sleep(config["rate_limit_delay"])

    result_df = pd.DataFrame(results).sort_values("index")

    try:
        output_path = os.path.abspath(config["output_path"])
        output_dir = os.path.dirname(output_path)

        if not os.path.exists(output_dir):
            os.makedirs(output_dir, exist_ok=True)
            logger.info(f"Created output directory: {output_dir}")

        logger.info(f"Attempting to save to: {output_path}")

        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            result_df.to_excel(writer, index=False)

        logger.info(f"File successfully saved to: {output_path}")
        logger.info(f"File size: {os.path.getsize(output_path) / 1024:.1f} KB")

    except PermissionError as pe:
        logger.error(f"Permission denied: {str(pe)}")
        logger.info("Please close the open Excel file and retry")
        csv_path = output_path.replace('.xlsx', '.csv')
        result_df.to_csv(csv_path)
        logger.info(f"Generated CSV backup file: {csv_path}")

    except Exception as e:
        logger.error(f"Save failed: {str(e)}")
        csv_path = output_path.replace('.xlsx', '.csv')
        result_df.to_csv(csv_path)
        logger.info(f"Generated CSV backup file: {csv_path}")

    # Generate Analysis Report
    success_count = result_df[result_df['status'] == 'success'].shape[0]
    format_warnings = result_df['model_output'].str.contains("Format Warning").sum()

    logger.info("\nAnalysis Report:")
    logger.info(f"Success Rate: {success_count / len(df):.1%}")
    logger.info(f"Format Warning Cases: {format_warnings}")
    logger.info(f"Average Attempt Count: {result_df['attempt_count'].mean():.1f}")

if __name__ == "__main__":
    main()
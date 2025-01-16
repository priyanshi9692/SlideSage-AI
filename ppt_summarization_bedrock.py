import os
import boto3
import json
import logging
from slidesage.ppt_data_preprocessing import PPTDataPreprocessing
from reporting_enginev2.template_path_retrieval.path_retrieval import ppt_template_path, \
 pandas_presentation_directory
LOGGER = logging.getLogger(__name__)


def initialize_bedrock_client(region_name="us-west-2"):
    """
    Initialize the Bedrock runtime client.

    Args:
        region_name (str): The AWS region for the Bedrock client. Defaults to 'us-west-2'.

    Returns:
        boto3.client: An initialized Bedrock client.
    """
    try:
        session = boto3.Session()
        bedrock = session.client('bedrock-runtime', region_name=region_name)
        return bedrock
    except Exception as e:
        raise RuntimeError(f"Error initializing Bedrock client: {str(e)}")


def query_bedrock_model(bedrock_client, model_id, prompt, max_token_count=2048):
    """
    Query the Bedrock model to generate a response based on a given prompt.

    Args:
        bedrock_client: Initialized Bedrock client.
        model_id (str): ID of the Bedrock model to query.
        prompt (str): Input text for the model.
        max_token_count (int): Maximum tokens for the response. Defaults to 2048.

    Returns:
        list: A list of output texts from the model.
    """
    try:
        results = []
        response = bedrock_client.invoke_model(
            modelId=model_id,
            contentType="application/json",
            accept="application/json",
            body=json.dumps({
                "inputText": prompt,
                "textGenerationConfig": {
                    "maxTokenCount": max_token_count,
                    "stopSequences": [],
                    "temperature": 0.7,
                    "topP": 1
                }
            })
        )
        response_body = json.loads(response['body'].read().decode('utf-8'))
        for result in response_body.get('results', []):
            print(f"Token count: {result.get('tokenCount')}")
            print(f"Output text: {result.get('outputText')}")
            print(f"Completion reason: {result.get('completionReason')}")
            results.append(result.get('outputText'))
        return results
    except Exception as e:
        raise RuntimeError(f"Error querying Bedrock model: {str(e)}")


def split_prompt(data, max_chars=3000):
    """
    Split the prompt into smaller chunks based on character count.
    """
    chunks = []
    current_chunk = ""

    for entry in data:
        entry_text = f"Slide Title: {entry['text']}\n"
        for table in entry.get('tables', []):
            for row in table:
                entry_text += ", ".join(row) + "\n"
        entry_text += "\n"

        # Check if adding this entry exceeds max_chars
        if len(current_chunk) + len(entry_text) > max_chars:
            chunks.append(current_chunk)
            current_chunk = entry_text
        else:
            current_chunk += entry_text

    # Add any remaining content
    if current_chunk:
        chunks.append(current_chunk)

    return chunks


# Generate prompt for insights
def generate_prompt(data):
    """
    Generate a natural language prompt for summarizing deliverability data.

    Args:
        data (list): List of dictionaries containing slide text and table data.

    Returns:
        str: A formatted natural language prompt.
    """
    try:
        prompt = "Analyze the following email deliverability data and provide a detailed, natural-language summary:\n\n"
        for entry in data:
            prompt += f"Slide Title: {entry['text']}\n"
            if entry.get('tables'):
                prompt += "Here is the tabular data:\n"
                for table in entry['tables']:
                    headers = ", ".join(table[0])
                    rows = "\n".join([", ".join(row) for row in table[1:]])
                    prompt += f"Headers: {headers}\nRows:\n{rows}\n"
            prompt += "\n"
        prompt += (
            "Based on the data provided, generate a detailed and easy-to-read summary. "
            "Explain deliverability performance, engagement metrics, and patterns in natural "
            "language without repeating the tabular data verbatim."
        )
        return prompt
    except Exception as e:
        raise RuntimeError(f"Error generating prompt: {str(e)}")


# Main function
def summarize_with_bedrock_titan(report_id):
    # Preprocessed slide deck data
    try:
        # Locate the PowerPoint file
        file_to_get_summarization = next(
            (os.path.join(pandas_presentation_directory, report_id, file)
             for file in os.listdir(os.path.join(pandas_presentation_directory, report_id))
             if file.endswith(".pptx")), None
        )
        if not file_to_get_summarization:
            LOGGER.error(f"No PowerPoint file found for report ID: {report_id}")
            return None

        # Process the PowerPoint file
        ppt_processor = PPTDataPreprocessing(file_to_get_summarization)
        slide_data = ppt_processor.preprocess_ppt()
        if not slide_data:
            LOGGER.error("Failed to preprocess content from the presentation.")
            return None

        LOGGER.info("\n--- Extracted Content ---\n")
        LOGGER.info(slide_data)

        # Initializing Bedrock client
        bedrock_client = initialize_bedrock_client()

        prompt = generate_prompt(slide_data)
        # Choosing the model titan-text-express-v1
        model_id = "amazon.titan-text-express-v1"
        response = query_bedrock_model(bedrock_client, model_id, prompt, max_token_count=4096)
        LOGGER.info(f"Generated Insights: {response}")

    except Exception as e:
        LOGGER.error(f"Error querying Bedrock model: {str(e)}")

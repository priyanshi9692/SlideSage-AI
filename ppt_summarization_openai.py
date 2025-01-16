import os
import openai
import logging
from config.config import Config
from reporting_enginev2.template_path_retrieval.path_retrieval import ppt_template_path, \
 pandas_presentation_directory
from slidesage.ppt_data_preprocessing import PPTDataPreprocessing

LOGGER = logging.getLogger(__name__)

AZURE_OPENAI_DEPLOYMENT = Config.AZURE_OPENAI_DEPLOYMENT
openai.api_type = Config.OPENAI_API_TYPE
openai.api_base = Config.AZURE_OPENAI_ENDPOINT
openai.api_version = Config.OPENAI_API_VERSION
openai.api_key = Config.OPENAI_API_KEY


def generate_prompt(slide_data):
    """
    Generates a natural language prompt for summarizing presentation data.

    This method constructs a detailed prompt by iterating through the slide data,
    which includes slide titles, content, and tables. The prompt is formatted to
    assist a language model in generating a coherent, natural-language summary of
    the presentation. The final prompt includes instructions for extracting key
    points, actionable insights, and important trends.

    Args:
        slide_data (list): A list of dictionaries, where each dictionary represents
            a slide and contains the following keys:
            - 'title' (str): The title of the slide.
            - 'content' (str): The textual content of the slide.
            - 'tables' (list): A list of tables on the slide, each represented as a
              list of rows, with the first row being the header.

    Returns:
        str: A detailed natural language prompt for summarization.
    """
    try:
        prompt = "You are a report summarizer. Please provide a detailed and easy-to-read summary of the following presentation:\n\n"
        for slide in slide_data:
            prompt += f"Slide Title: {slide['title']}\n"
            if slide['content']:
                prompt += f"Slide Content: {slide['content']}\n"
            if slide['tables']:
                prompt += "Slide Tables:\n"
                for table in slide['tables']:
                    headers = ", ".join(table[0])
                    rows = "\n".join([", ".join(row) for row in table[1:]])
                    prompt += f"Headers: {headers}\nRows:\n{rows}\n"
            prompt += "\n"
        prompt += (
            "Summarize the key points from all slides in a coherent and natural-language manner. "
            "Provide actionable insights and highlight important trends or patterns."
        )
        return prompt
    except Exception as e:
        LOGGER.error(f"Error generating prompt: {e}")
        return None


def summarize_with_azure_openai(prompt):
    """
    Generates a summary using Azure OpenAI ChatCompletion API based on the provided prompt.

    This method sends a prompt to the Azure OpenAI service, where it is processed to create a
    natural language summary. The summary is returned in a formatted string. In case of any
    errors during the API call or response processing, the method logs the error and returns None.

    Args:
        prompt (str): The input text or query to be summarized by the Azure OpenAI service.

    Returns:
        str: The formatted summary generated by the model if successful.
        None: If an error occurs during the summarization process.
    """
    try:
        response = openai.ChatCompletion.create(
            engine=AZURE_OPENAI_DEPLOYMENT,
            messages=[
                {"role": "system", "content": "You are an expert data analyst and report summarizer."},
                {"role": "user", "content": prompt}
            ],
            temperature=0.7
        )
        summary = response['choices'][0]['message']['content']
        formatted_summary = f"""{summary}"""
        return formatted_summary
    except Exception as e:
        LOGGER.error(f"Error: {e}")
        return None


def ppt_summarization(report_id):
    """
    Processes a PowerPoint presentation, extracts its content, and generates a natural language summary.

    Steps:
    1. Locates the PowerPoint (.pptx) file in the specified directory using the given `report_id`.
    2. Extracts slide text and table data from the presentation using `PPTDataPreprocessing`.
    3. Generates a prompt based on the extracted content.
    4. Uses Azure OpenAI to generate a summary based on the prompt.
    5. Logs the extracted content, generated prompt, and resulting summary.

    Args:
        report_id (str): The identifier for the report directory containing the PowerPoint file.

    Returns:
        str: The generated summary if successful; otherwise, logs an error and returns None.
    """
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

        # Generating a prompt based on the extracted content
        prompt = generate_prompt(slide_data)
        LOGGER.info(f"Generated Prompt:\n{prompt[:100]}...\n")

        # Using Azure OpenAI to generate the summary
        summary = summarize_with_azure_openai(prompt)
        if summary:
            LOGGER.info(f"Generated Summary:\n {summary}")
            return summary
        else:
            LOGGER.error("Failed to generate summary.")
            return None

    except Exception as e:
        LOGGER.error(f"An unexpected error occurred: {e}")
        return None

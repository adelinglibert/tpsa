**Simple Cybersecurity Third Party Risk Assessment Generator
**

This project provides a script to generate a comprehensive third party risk assessment document. The script combines multiple DOCX and PDF files into a single PDF, analyzes the combined PDF using the OpenAI GPT-4 model, and creates a formatted memorandum based on a provided template.

Features

-Combines DOCX and PDF files into a single PDF.
-Analyzes the combined PDF using the OpenAI GPT-4 model.
-Generates a detailed cybersecurity third party risk assessment.
-Includes sources for verifying the quality of the LLM output and for further investigation.

Prerequisites

    Python 3.7 or higher
    OpenAI API key
    Required Python packages (see below)


Clone the repository:
    
    git clone https://github.com/adelinglibert/tpsa.git
    cd tpsa

Install the required Python packages.

Replace 'your_openai_api_key' in the script with your actual OpenAI API key.

Usage

    Prepare your documents:

    Place the DOCX and PDF files you want to analyze in a folder.

    Prepare your template:

    Ensure you have a .docx template file with placeholders for the analysis results.

    Run the script:
    bash

    python main.py <folder_path> <template_path> <output_path>

        <folder_path>
        : The path to the folder containing the files to analyze.
        <template_path>
        : The path to the .docx template file.
        <output_path>
        : The path to save the generated memorandum.

Example

    python main.py ./documents ./template/template.docx ./output/memorandum.docx

Prompt for OpenAI API

The prompt used for analysis is designed to instruct the GPT-4 model to provide a detailed cybersecurity risk assessment. It includes sections for services, technical setup, security measures, data flows, third parties, and risks and recommendations.

Contributing

Contributions are welcome! Please open an issue or submit a pull request for any improvements or bug fixes.

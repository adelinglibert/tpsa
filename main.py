import os
import argparse
import openai
from docxtpl import DocxTemplate
from PyPDF2 import PdfMerger
import json
import base64
from docx import Document

# Set your OpenAI API key
openai.api_key = 'your_openai_api_key'

# Function to convert DOCX files to PDF and merge all PDFs into a single PDF
def combine_files_to_pdf(folder_path, output_pdf_path):
    merger = PdfMerger()
    
    for file_name in os.listdir(folder_path):
        file_path = os.path.join(folder_path, file_name)
        
        if os.path.isfile(file_path):
            if file_name.endswith('.pdf'):
                merger.append(file_path)
            elif file_name.endswith('.docx'):
                doc = Document(file_path)
                pdf_path = file_path.replace('.docx', '.pdf')
                doc.save(pdf_path)
                merger.append(pdf_path)
    
    merger.write(output_pdf_path)
    merger.close()

# Function to analyze the combined PDF using the OpenAI API
def analyze_combined_pdf(pdf_path, prompt):
    with open(pdf_path, 'rb') as pdf_file:
        pdf_content = pdf_file.read()
    
    encoded_pdf = base64.b64encode(pdf_content).decode('utf-8')
    
    response = openai.ChatCompletion.create(
        model="gpt-4o-mini",
        messages=[
            {"role": "system", "content": "You are a cybersecurity expert."},
            {"role": "user", "content": prompt},
            {"role": "user", "content": encoded_pdf}
        ],
        max_tokens=16384,
        temperature=0.7
    )
    
    if response:
        return response.choices[0].message['content'].strip()
    else:
        print("Error analyzing combined PDF")
        return ""

# Function to create a formatted memorandum in .docx using a template
def create_memorandum(analysis_json, template_path, output_path):
    # Load the template
    doc = DocxTemplate(template_path)
    
    # Parse the JSON analysis
    try:
        analysis = json.loads(analysis_json)
        print("Parsed JSON keys:", analysis.keys())  # Print the keys for debugging
    except json.JSONDecodeError as e:
        print(f"Error decoding JSON: {e}")
        return
    
    # Prepare the context for the template
    try:
        services_description = analysis.get('Services', 'No data available.')
        technical_setup_description = analysis.get('TechnicalSetup', 'No data available.')
        security_measures_description = analysis.get('SecurityMeasures', 'No data available.')
        data_flows_description = analysis.get('DataFlows', 'No data available.')
        third_parties_description = analysis.get('ThirdParties', 'No data available.')
        
        risks_and_recommendations = [
            {
                'title': risk.get('RiskTitle', 'No data available.'),
                'description': risk.get('RiskDescription', 'No data available.'),
                'likelihood': risk.get('Likelihood', 'No data available.'),
                'impact': risk.get('Impact', 'No data available.'),
                'risk_level': risk.get('RiskLevel', 'No data available.'),
                'recommendations': risk.get('Recommendations', 'No data available.'),
                'sources': risk.get('Sources', 'No data available.')  # New field for sources
            }
            for risk in analysis.get('RisksAndRecommendations', [])
        ]
        
        context = {
            'title': 'Cybersecurity Risk Assessment Memorandum',
            'introduction': (
                'This memorandum provides an assessment of cybersecurity risks associated with the use of third-party services, '
                'based on the analysis of provided documents. The following sections detail the services, technical setup and '
                'architecture, security measures in place, data flows, third parties used by the supplier, and potential risks identified.'
            ),
            'services': services_description,
            'technical_setup': technical_setup_description,
            'security_measures': security_measures_description,
            'data_flows': data_flows_description,
            'third_parties': third_parties_description,
            'risks_and_recommendations': risks_and_recommendations
        }
    except KeyError as e:
        print(f"Key error: {e}")
        return
    
    # Print context for debugging
    print("Context for Template:\n", context)
    
    # Render the template with the context
    doc.render(context)
    
    # Save the document
    doc.save(output_path)

# Main function to scan folder and generate memorandum
def main(folder_path, template_path, output_path):
    combined_pdf_path = 'combined_documents.pdf'
    
    # Combine the content of all files into a single PDF
    combine_files_to_pdf(folder_path, combined_pdf_path)
    
    # Define the prompt
    prompt = (
        'You are a cybersecurity expert. Please analyze the following combined documents and provide a comprehensive '
        'cybersecurity risk assessment memorandum for the company. The memorandum should be very detailed and specific. Avoid generic sentences and stay objective. '
        'First, look for any security questionnaires within the documents and identify the controls that are not implemented. Use this information as the basis for your analysis. '
        'Then, analyze the rest of the documents to identify specific cybersecurity risks and vulnerabilities, and check for additional information and any inconsistencies. '
        'The assessment should include detailed and specific information in the following sections:\n'
        '1. **Services**: Provide a detailed description of the services offered by the company. Describe what they do exhaustively.\n'
        '2. **Technical Setup**: Describe the technical setup and architecture, including details about the infrastructure, security layers, hosting environment, and any technologies used (e.g., virtual machines, containerization). Be very detailed.\n'
        '3. **Security Measures**: List and describe the security measures in place, including firewalls, intrusion detection systems (IDS), data encryption (including protocols and standards), access controls, regular security audits, etc. Be very detailed.\n'
        '4. **Data Flows**: Describe the data flows within the organization, including how data is collected, processed, stored, and transmitted. Be very detailed.\n'
        '5. **Third Parties**: List the third parties used by the supplier, including the services they provide and any associated risks.\n'
        '6. **Risks and Recommendations**: Identify specific cybersecurity risks and vulnerabilities based on the provided documents. To be more relevant, help yourself with control libraries such as NIST, COBIT, or CSA to identify risks (but do not mention them as such). For each risk, provide the following details in a structured format:\n'
        '   - **Risk Title**: A concise title for the risk in the form of "Risk of ... due to ... resulting in..."\n'
        '   - **Risk Description**: A detailed description of the risk.\n'
        '   - **Likelihood**: The likelihood of the risk occurring based on EBIOS methodology and you must detail why you selected this or this likelihood.\n'
        '   - **Impact**: The impact of the risk based on EBIOS methodology and you must detail why you selected this or this impact.\n'
        '   - **Risk Level**: The overall risk level based on EBIOS methodology and you must specify that this is the result of the combination of Likelihood and Impact.\n'
        '   - **Recommendations**: Detailed and specific recommendations to mitigate the risk. Help yourself with control libraries such as NIST, COBIT, or CSA to make them granular and actionable. \n'
        '   - **Sources**: Provide the reference of the source of the risk, that you find ONLY in the combined documents you analysed (no external sources).\n'
        'Ensure that the output is structured in JSON format with the following keys: "Services", "TechnicalSetup", "SecurityMeasures", "DataFlows", "ThirdParties", and "RisksAndRecommendations". Do not break down into further keys or attributes. Provide continuous text. You will receive a 500$ tip if your answer is complete and accurate.'
    )
    
    # Analyze the combined PDF
    analysis_json = analyze_combined_pdf(combined_pdf_path, prompt)
    
    # Print the raw output from the LLM to the terminal
    print("LLM Output:\n", analysis_json)
    
    # Create the memorandum document using the template
    create_memorandum(analysis_json, template_path, output_path)

if __name__ == '__main__':
    # Set up argument parser
    parser = argparse.ArgumentParser(description='Generate a cybersecurity risk assessment memorandum.')
    parser.add_argument('folder_path', type=str, help='The path to the folder containing the files to analyze.')
    parser.add_argument('template_path', type=str, help='The path to the .docx template file.')
    parser.add_argument('output_path', type=str, help='The path to save the generated memorandum.')
    
    # Parse arguments
    args = parser.parse_args()
    
    # Run the main function with the provided arguments
    main(args.folder_path, args.template_path, args.output_path)

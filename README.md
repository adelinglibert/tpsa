# TPSA – Third-Party Security Assessment Generator

A CLI tool that analyses PDF and DOCX supplier documents and generates a comprehensive **cybersecurity risk assessment memorandum** (`.docx`) — including a **post-assessment remediation plan**.

Supports any major LLM provider: **OpenAI, Anthropic (Claude), Google Gemini, Mistral**.

---

## Features

- **LLM-agnostic** — switch providers and models via a single flag; credentials from env vars or `--api-key`
- **Accurate document ingestion** — text extracted from PDFs and DOCX files (no lossy base64 encoding)
- **EBIOS RM–aligned risk scoring** — explicit 4-level likelihood and impact scale with justifications
- **M&A due diligence mode** — extended prompts covering cyber debt, integration risks, and hidden liabilities
- **Regulatory alignment** — optional coverage of NIS2, DORA, SOC 2, ISO 27001, GDPR
- **Post-assessment remediation plan** — priority actions (0–30 days), medium-term (1–6 months), strategic (6–18 months), governance recommendations, and ownership matrix
- **Raw JSON export** — optionally save the LLM output for further processing

---

## Prerequisites

- Python 3.9 or higher
- API key for at least one supported LLM provider

---

## Installation

```bash
git clone https://github.com/adelinglibert/tpsa.git
cd tpsa
pip install -r requirements.txt
```

You only need to install the provider SDK(s) you intend to use.  
For example, if you only use OpenAI:

```bash
pip install openai pdfplumber python-docx docxtpl PyPDF2
```

---

## API Keys

Set the environment variable for your chosen provider — **never hardcode keys in the script**:

| Provider  | Environment Variable  |
|-----------|-----------------------|
| OpenAI    | `OPENAI_API_KEY`      |
| Anthropic | `ANTHROPIC_API_KEY`   |
| Gemini    | `GEMINI_API_KEY`      |
| Mistral   | `MISTRAL_API_KEY`     |

Or pass the key directly with `--api-key`.

---

## Usage

```
python main.py <folder_path> <template_path> <output_path> [options]
```

### Arguments

| Argument         | Description                                              |
|------------------|----------------------------------------------------------|
| `folder_path`    | Folder containing the PDF/DOCX documents to analyse     |
| `template_path`  | DOCX template file (with Jinja2-style placeholders)     |
| `output_path`    | Where to save the generated memorandum                  |

### Options

| Option                       | Description                                                                            | Default             |
|------------------------------|----------------------------------------------------------------------------------------|---------------------|
| `--provider`, `-p`           | LLM provider: `openai`, `anthropic`, `gemini`, `mistral`                               | `openai`            |
| `--model`, `-m`              | Model name (see defaults below)                                                        | Provider default    |
| `--api-key`, `-k`            | API key (falls back to env var)                                                        | –                   |
| `--context`, `-c`            | Assessment context: `tprm` (standard) or `ma` (M&A due diligence)                     | `tprm`              |
| `--regulations`, `-r`        | Comma-separated frameworks: `nis2`, `dora`, `soc2`, `iso27001`, `gdpr`                | –                   |
| `--save-raw`                 | Save raw LLM JSON output alongside the DOCX                                           | off                 |

### Default models per provider

| Provider  | Default model            |
|-----------|--------------------------|
| OpenAI    | `gpt-4o`                 |
| Anthropic | `claude-opus-4-5`        |
| Gemini    | `gemini-1.5-pro`         |
| Mistral   | `mistral-large-latest`   |

---

## Examples

**Standard TPRM with OpenAI GPT-4o:**
```bash
export OPENAI_API_KEY=sk-...
python main.py ./documents ./template.docx ./report.docx
```

**M&A due diligence with Anthropic Claude:**
```bash
export ANTHROPIC_API_KEY=sk-ant-...
python main.py ./documents ./template.docx ./report_ma.docx \
  --provider anthropic --model claude-opus-4-5 \
  --context ma --regulations nis2,dora,gdpr
```

**Google Gemini, saving raw JSON:**
```bash
export GEMINI_API_KEY=AIza...
python main.py ./documents ./template.docx ./report.docx \
  --provider gemini --save-raw
```

**Mistral with inline key:**
```bash
python main.py ./documents ./template.docx ./report.docx \
  --provider mistral --api-key <your-key> --regulations soc2,iso27001
```

---

## Output Sections

The generated memorandum contains:

| Section                       | Description                                                       |
|-------------------------------|-------------------------------------------------------------------|
| Services                      | What the supplier does                                            |
| Technical Setup               | Infrastructure, hosting, architecture                             |
| Security Measures             | Controls in place (firewalls, IAM, encryption, audits…)           |
| Data Flows                    | How data is collected, processed, stored, and transmitted         |
| Third Parties                 | Sub-processors and their associated risks                         |
| Risks & Recommendations       | EBIOS RM–scored risk register with granular mitigations           |
| **Remediation Plan**          | **Prioritised roadmap: 0–30 days / 1–6 months / 6–18 months**    |
| **Governance Recommendations**| **Contractual clauses, right-to-audit, monitoring cadence**       |
| **Ownership Matrix**          | **RACI-style mapping of risks to CISO, DPO, Legal, Procurement**  |

## Contributing

Contributions are welcome. Please open an issue or submit a pull request.

## License

See [LICENSE](LICENSE).

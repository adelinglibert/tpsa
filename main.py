"""
TPSA – Third-Party Security Assessment Generator
=================================================
LLM-agnostic CLI tool for generating comprehensive cybersecurity risk
assessment memoranda, post-assessment remediation plans, IT stack
comparisons, remediation effort estimates, and executive scorecards.

Supported LLM providers:
    openai     – OpenAI (GPT-4o, GPT-4o-mini, etc.)
    anthropic  – Anthropic Claude (claude-opus-4-5, claude-3-5-sonnet, etc.)
    gemini     – Google Gemini (gemini-1.5-pro, gemini-2.0-flash, etc.)
    mistral    – Mistral AI (mistral-large-latest, etc.)

Usage:
    python main.py <folder_path> <template_path> <output_path> [options]

Examples:
    python main.py ./docs ./template.docx ./report.docx --provider openai --model gpt-4o
    python main.py ./docs ./template.docx ./report.docx --provider anthropic --model claude-opus-4-5
    python main.py ./docs ./template.docx ./report.docx --provider gemini
    python main.py ./docs ./template.docx ./report.docx --provider mistral --context ma --regulations nis2,dora
    python main.py ./docs ./template.docx ./report.docx --effort --org-size Mid
    python main.py ./docs ./template.docx ./report.docx --stack ./acquirer_stack.json
    python main.py ./docs ./template.docx ./report.docx --scorecard --baseline ./previous_report.json
"""

from __future__ import annotations

import argparse
import json
import os
import sys
import textwrap
from enum import Enum
from pathlib import Path
from typing import Optional

import pdfplumber
from docx import Document
from docxtpl import DocxTemplate
from PyPDF2 import PdfMerger

# ---------------------------------------------------------------------------
# Provider enum & defaults
# ---------------------------------------------------------------------------

class Provider(str, Enum):
    OPENAI = "openai"
    ANTHROPIC = "anthropic"
    GEMINI = "gemini"
    MISTRAL = "mistral"


DEFAULT_MODELS: dict[Provider, str] = {
    Provider.OPENAI: "gpt-4o",
    Provider.ANTHROPIC: "claude-opus-4-5",
    Provider.GEMINI: "gemini-1.5-pro",
    Provider.MISTRAL: "mistral-large-latest",
}

ENV_KEY_NAMES: dict[Provider, str] = {
    Provider.OPENAI: "OPENAI_API_KEY",
    Provider.ANTHROPIC: "ANTHROPIC_API_KEY",
    Provider.GEMINI: "GEMINI_API_KEY",
    Provider.MISTRAL: "MISTRAL_API_KEY",
}

# Max characters of extracted text to send in a single LLM call.
# Roughly ~120 000 tokens at 4 chars/token → safe for 128 k-context models.
MAX_TEXT_CHARS = 480_000

# ---------------------------------------------------------------------------
# Document text extraction
# ---------------------------------------------------------------------------

def extract_text_from_pdf(pdf_path: str) -> str:
    """Extract text from a PDF using pdfplumber (structure-aware)."""
    pages = []
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for i, page in enumerate(pdf.pages, 1):
                text = page.extract_text() or ""
                if text.strip():
                    pages.append(f"--- Page {i} ---\n{text}")
    except Exception as exc:
        print(f"  [WARN] Could not extract text from {pdf_path}: {exc}")
    return "\n\n".join(pages)


def extract_text_from_docx(docx_path: str) -> str:
    """Extract text from a DOCX file paragraph by paragraph."""
    try:
        doc = Document(docx_path)
        paragraphs = [p.text for p in doc.paragraphs if p.text.strip()]
        # Also extract tables
        for table in doc.tables:
            for row in table.rows:
                row_text = " | ".join(cell.text.strip() for cell in row.cells if cell.text.strip())
                if row_text:
                    paragraphs.append(row_text)
        return "\n".join(paragraphs)
    except Exception as exc:
        print(f"  [WARN] Could not extract text from {docx_path}: {exc}")
        return ""


def collect_text_from_folder(folder_path: str) -> str:
    """
    Walk a folder, extract text from all PDF and DOCX files, and concatenate
    them with clear file-level headings.
    """
    folder = Path(folder_path)
    if not folder.is_dir():
        sys.exit(f"[ERROR] Folder not found: {folder_path}")

    sections: list[str] = []
    supported = {".pdf", ".docx"}
    files = sorted(f for f in folder.rglob("*") if f.suffix.lower() in supported)

    if not files:
        sys.exit(f"[ERROR] No PDF or DOCX files found in: {folder_path}")

    print(f"\n[INFO] Found {len(files)} document(s) to analyse:")
    for f in files:
        print(f"  • {f.name}")
        if f.suffix.lower() == ".pdf":
            text = extract_text_from_pdf(str(f))
        else:
            text = extract_text_from_docx(str(f))

        if text.strip():
            sections.append(f"=== FILE: {f.name} ===\n\n{text}")
        else:
            print(f"  [WARN] No extractable text found in {f.name} — skipping.")

    combined = "\n\n".join(sections)

    if len(combined) > MAX_TEXT_CHARS:
        print(
            f"[WARN] Combined text is {len(combined):,} chars — truncating to "
            f"{MAX_TEXT_CHARS:,} to stay within context limits."
        )
        combined = combined[:MAX_TEXT_CHARS]

    return combined


# ---------------------------------------------------------------------------
# LLM provider adapters
# ---------------------------------------------------------------------------

def _resolve_api_key(provider: Provider, cli_key: Optional[str]) -> str:
    key = cli_key or os.environ.get(ENV_KEY_NAMES[provider], "")
    if not key:
        sys.exit(
            f"[ERROR] No API key found for provider '{provider.value}'.\n"
            f"  Set the {ENV_KEY_NAMES[provider]} environment variable "
            f"or pass --api-key."
        )
    return key


def call_llm_openai(system_prompt: str, user_prompt: str, model: str, api_key: str) -> str:
    try:
        from openai import OpenAI
    except ImportError:
        sys.exit("[ERROR] openai package not installed. Run: pip install openai>=1.0.0")

    client = OpenAI(api_key=api_key)
    response = client.chat.completions.create(
        model=model,
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt},
        ],
        max_tokens=16384,
        temperature=0.3,
        response_format={"type": "json_object"},
    )
    return response.choices[0].message.content.strip()


def call_llm_anthropic(system_prompt: str, user_prompt: str, model: str, api_key: str) -> str:
    try:
        import anthropic
    except ImportError:
        sys.exit("[ERROR] anthropic package not installed. Run: pip install anthropic>=0.25.0")

    client = anthropic.Anthropic(api_key=api_key)
    message = client.messages.create(
        model=model,
        max_tokens=16000,
        system=system_prompt,
        messages=[{"role": "user", "content": user_prompt}],
    )
    return message.content[0].text.strip()


def call_llm_gemini(system_prompt: str, user_prompt: str, model: str, api_key: str) -> str:
    try:
        import google.generativeai as genai
    except ImportError:
        sys.exit("[ERROR] google-generativeai package not installed. Run: pip install google-generativeai>=0.5.0")

    genai.configure(api_key=api_key)
    gemini_model = genai.GenerativeModel(
        model_name=model,
        system_instruction=system_prompt,
    )
    response = gemini_model.generate_content(user_prompt)
    return response.text.strip()


def call_llm_mistral(system_prompt: str, user_prompt: str, model: str, api_key: str) -> str:
    try:
        from mistralai import Mistral
    except ImportError:
        sys.exit("[ERROR] mistralai package not installed. Run: pip install mistralai>=0.4.0")

    client = Mistral(api_key=api_key)
    response = client.chat.complete(
        model=model,
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt},
        ],
        max_tokens=16384,
        temperature=0.3,
    )
    return response.choices[0].message.content.strip()


def call_llm(
    system_prompt: str,
    user_prompt: str,
    provider: Provider,
    model: str,
    api_key: str,
) -> str:
    """Unified LLM dispatch."""
    dispatch = {
        Provider.OPENAI: call_llm_openai,
        Provider.ANTHROPIC: call_llm_anthropic,
        Provider.GEMINI: call_llm_gemini,
        Provider.MISTRAL: call_llm_mistral,
    }
    print(f"\n[INFO] Calling {provider.value} / {model} …")
    return dispatch[provider](system_prompt, user_prompt, model, api_key)


# ---------------------------------------------------------------------------
# Prompt construction
# ---------------------------------------------------------------------------

REGULATION_DESCRIPTIONS: dict[str, str] = {
    "nis2": (
        "EU NIS2 Directive (ensuring the supplier meets obligations for network and "
        "information system security, incident reporting, and supply chain risk management)"
    ),
    "dora": (
        "EU DORA (Digital Operational Resilience Act – ICT risk management, incident "
        "classification, resilience testing, and third-party ICT provider oversight)"
    ),
    "soc2": (
        "SOC 2 Trust Services Criteria (security, availability, processing integrity, "
        "confidentiality, and privacy)"
    ),
    "iso27001": (
        "ISO/IEC 27001 (information security management system requirements and Annex A controls)"
    ),
    "gdpr": (
        "EU GDPR (data processing lawfulness, data subject rights, data transfers, "
        "and processor obligations)"
    ),
}

SYSTEM_PROMPT = textwrap.dedent("""
    You are a senior cybersecurity consultant specialising in third-party risk management (TPRM),
    mergers & acquisitions (M&A) due diligence, and supply chain security. You have deep expertise
    in the EBIOS Risk Manager methodology, NIST CSF, ISO 27001/27005, SOC 2, NIS2, and DORA.

    Your role is to produce rigorous, evidence-based, and actionable security assessments.
    Never produce generic boilerplate. Every claim must be traceable to the documents provided.
    Use precise, professional language appropriate for a C-suite / board-level audience.
    Output ONLY valid JSON — no markdown fences, no commentary outside the JSON structure.
""").strip()


# ---------------------------------------------------------------------------
# Org-size effort scaling tables (Feature 2)
# ---------------------------------------------------------------------------

ORG_SIZE_LABELS = ["SME", "Mid", "Large", "Enterprise"]

# effort_factor: multiplier applied to base man-day estimates per action
ORG_SIZE_PROFILES: dict[str, dict] = {
    "SME": {
        "label": "SME (<250 employees)",
        "effort_factor": 1.0,
        "note": "Lean security team; expect 1 FTE security generalist. Actions may require external consultant support.",
    },
    "Mid": {
        "label": "Mid-size (250–2 500 employees)",
        "effort_factor": 1.4,
        "note": "Dedicated security team of 2–5. Most actions are in-house feasible with occasional specialist uplift.",
    },
    "Large": {
        "label": "Large (2 500–10 000 employees)",
        "effort_factor": 2.0,
        "note": "Structured security function. Governance, change management, and stakeholder alignment add overhead.",
    },
    "Enterprise": {
        "label": "Enterprise (>10 000 employees)",
        "effort_factor": 3.0,
        "note": "Complex multi-team environments. Procurement cycles, compliance reviews, and CAB add significant lead time.",
    },
}


def build_user_prompt(
    combined_text: str,
    context_mode: str,
    regulations: list[str],
    enable_effort: bool = False,
    org_size: str = "Mid",
    acquirer_stack: Optional[dict] = None,
    enable_scorecard: bool = False,
    baseline: Optional[dict] = None,
    # Features A–E
    enable_contracts: bool = False,
    enable_red_flag: bool = False,
    enable_supply_chain: bool = False,
    enable_questionnaire: bool = False,
    enable_heatmap: bool = False,
) -> str:
    context_clause = ""
    if context_mode == "ma":
        context_clause = (
            "\n\nCONTEXT: This assessment is performed as part of a Merger & Acquisition (M&A) "
            "cybersecurity due diligence exercise. In addition to standard TPRM risks, evaluate "
            "integration risks, hidden liabilities, cyber debt (technical security debt), "
            "potential regulatory exposure from acquisition, and any indicators of past breaches "
            "or ongoing vulnerabilities that could affect deal valuation or post-merger operations."
        )

    reg_clause = ""
    if regulations:
        reg_lines = "\n".join(
            f"  - {REGULATION_DESCRIPTIONS.get(r.lower(), r)}"
            for r in regulations
        )
        reg_clause = (
            f"\n\nREGULATORY SCOPE: In addition to general risk identification, explicitly "
            f"assess compliance gaps against the following frameworks/regulations:\n{reg_lines}"
        )

    ebios_scale = textwrap.dedent("""
        EBIOS RM SCORING SCALE (use these exact labels):
          Likelihood: 1-Minimal, 2-Significant, 3-Strong, 4-Maximum
          Impact:     1-Negligible, 2-Limited, 3-Important, 4-Critical
          Risk Level: Combine Likelihood × Impact following the EBIOS RM matrix:
            - 1×1..2×1..1×2 → Low
            - 2×2..3×1..1×3 → Medium
            - 3×2..2×3..4×1..1×4 → High
            - 3×3..4×2..2×4..3×4..4×3..4×4 → Critical
    """).strip()

    # ------------------------------------------------------------------
    # Build the JSON schema dynamically based on enabled features
    # ------------------------------------------------------------------

    # --- Core schema (always present) ---
    core_schema = textwrap.dedent("""
        {
          "CompanyName": "string – name of the company being assessed",
          "AssessmentDate": "string – today's date in YYYY-MM-DD",
          "Services": "string – exhaustive description of all services offered",
          "TechnicalSetup": "string – infrastructure, hosting, architecture, technologies",
          "SecurityMeasures": "string – controls in place: firewalls, IDS/IPS, encryption, IAM, audits",
          "DataFlows": "string – how data is collected, processed, stored, transmitted",
          "ThirdParties": "string – all sub-processors and third parties, with their roles and risks",
          "RisksAndRecommendations": [
            {
              "RiskTitle": "Risk of [X] due to [Y] resulting in [Z]",
              "RiskDescription": "string",
              "Likelihood": "string – EBIOS RM label + detailed justification from documents",
              "Impact": "string – EBIOS RM label + detailed justification from documents",
              "RiskLevel": "string – EBIOS RM level (Low/Medium/High/Critical) + explanation",
              "Recommendations": "string – actionable, granular mitigations referencing NIST/ISO/CIS controls",
              "Sources": "string – document name and relevant section/page from the provided documents only"
            }
          ],
          "RemediationPlan": {
            "PriorityActions": [
              {
                "Action": "string – specific action to take",
                "Rationale": "string – why this is urgent",
                "Owner": "string – suggested accountable role (e.g., CISO, DPO, Procurement)",
                "Deadline": "string – e.g., '0–30 days'"
              }
            ],
            "MediumTermActions": [
              {
                "Action": "string",
                "Rationale": "string",
                "Owner": "string",
                "Deadline": "string – e.g., '1–6 months'"
              }
            ],
            "StrategicActions": [
              {
                "Action": "string",
                "Rationale": "string",
                "Owner": "string",
                "Deadline": "string – e.g., '6–18 months'"
              }
            ],
            "GovernanceRecommendations": "string – contractual clauses, SLA/KPI requirements, right-to-audit, incident notification SLAs, continuous monitoring cadence",
            "OwnershipMatrix": "string – summary mapping of risks to accountable and responsible parties across CISO, DPO, Legal, Procurement, IT, and Business"
          }
    """).strip()

    # --- Feature 2: Effort Estimator ---
    effort_schema = ""
    effort_clause = ""
    if enable_effort:
        profile = ORG_SIZE_PROFILES.get(org_size, ORG_SIZE_PROFILES["Mid"])
        effort_clause = textwrap.dedent(f"""
            EFFORT ESTIMATION CONTEXT:
            The assessed organisation size is: {profile['label']}.
            {profile['note']}
            Effort factor vs. baseline SME: {profile['effort_factor']}×.
            Apply this scaling when estimating man-days.
        """).strip()

        effort_schema = textwrap.dedent("""
            ,
              "EffortEstimate": {
                "OrgSize": "string – organisation size tier used for calibration",
                "TotalEffortDaysMin": "integer – pessimistic total man-days across all remediations",
                "TotalEffortDaysMid": "integer – realistic total man-days",
                "TotalEffortDaysMax": "integer – optimistic total man-days",
                "ActionEstimates": [
                  {
                    "Action": "string – action description (match wording from RemediationPlan)",
                    "Phase": "string – Priority / MediumTerm / Strategic",
                    "EffortDaysMin": "integer",
                    "EffortDaysMid": "integer",
                    "EffortDaysMax": "integer",
                    "RoleProfile": "string – skill profile required (e.g. IAM Engineer, Security Architect, DPO, External Consultant)",
                    "InHouseFeasible": "boolean – true if typical in-house team can execute without specialist external help",
                    "Notes": "string – caveats, dependencies, or scaling assumptions"
                  }
                ],
                "KeyAssumptions": "string – main assumptions behind these estimates (team maturity, parallelisation, tooling availability)"
              }
        """).strip()

    # --- Feature 1: IT Stack Comparator ---
    stack_schema = ""
    stack_clause = ""
    if acquirer_stack:
        stack_json = json.dumps(acquirer_stack, indent=2)
        stack_clause = textwrap.dedent(f"""
            IT STACK COMPARISON:
            The acquirer / client's IT and security stack is provided below. Extract the supplier's
            stack from the documents and compare against this acquirer stack across the following
            layers: IAM, SIEM/SOC, ITSM, Cloud Platform, Network/Perimeter, Endpoint, Data Platform,
            and BCDR. Identify: (a) compatible tools already in use by both parties, (b) functional
            overlaps that may create redundancy or licensing conflicts post-integration, (c) gaps or
            incompatibilities that create integration hurdles, and (d) consolidation opportunities.

            ACQUIRER STACK:
            {stack_json}
        """).strip()

        stack_schema = textwrap.dedent("""
            ,
              "ITStackComparison": {
                "SupplierStack": {
                  "IAM": "string – identity and access management tools/protocols extracted from documents",
                  "SIEM": "string – SIEM / SOC tooling",
                  "ITSM": "string – IT service management platform",
                  "Cloud": "string – cloud provider(s) and services",
                  "Network": "string – firewalls, proxies, VPN, SD-WAN",
                  "Endpoint": "string – EDR / antivirus / MDM",
                  "DataPlatform": "string – databases, data warehouses, object storage",
                  "BCDR": "string – backup, DR, and business continuity tooling"
                },
                "CompatibilityMatrix": [
                  {
                    "Layer": "string – e.g. IAM",
                    "SupplierTool": "string",
                    "AcquirerTool": "string",
                    "CompatibilityStatus": "Compatible / Overlap / Incompatible / Unknown",
                    "IntegrationHurdle": "string – specific challenge or dependency",
                    "ConsolidationOpportunity": "string – recommended path forward"
                  }
                ],
                "OverallIntegrationRisk": "Low / Medium / High / Critical",
                "IntegrationSummary": "string – narrative summary of key integration hurdles and recommended consolidation path"
              }
        """).strip()

    # --- Feature 3: Executive Scorecard ---
    scorecard_schema = ""
    scorecard_clause = ""
    if enable_scorecard:
        baseline_str = ""
        if baseline:
            prev = baseline.get("ExecutiveScorecard", {})
            prev_domains = {d["Domain"]: d["Score"] for d in prev.get("Domains", [])}
            baseline_str = (
                "\nBASELINE SCORES FROM PREVIOUS ASSESSMENT (compare and compute delta):\n"
                + json.dumps(prev_domains, indent=2)
            )

        scorecard_clause = textwrap.dedent(f"""
            EXECUTIVE SCORECARD:
            Produce a single-page executive scorecard assessing maturity across the following
            eight security domains: Governance & Risk Management, Identity & Access Management,
            Data Protection & Privacy, Infrastructure & Network Security, Application Security,
            Incident Response & Recovery, Third-Party & Supply Chain Risk, Compliance & Audit.
            Score each domain: Red (significant gaps), Amber (partial implementation), Green
            (adequate controls). Be evidence-based; cite the relevant document finding.
            {baseline_str}
        """).strip()

        scorecard_schema = textwrap.dedent("""
            ,
              "ExecutiveScorecard": {
                "OverallRating": "Red / Amber / Green – single organisation-level verdict",
                "OverallSummary": "string – 2–3 sentence board-level summary",
                "Domains": [
                  {
                    "Domain": "string – domain name",
                    "Score": "Red / Amber / Green",
                    "Delta": "string – vs. previous assessment: Improved / Deteriorated / Unchanged / N/A",
                    "KeyFinding": "string – single most important finding for this domain",
                    "EvidenceSource": "string – document name + section"
                  }
                ]
              }
        """).strip()


    # --- Feature A: Contractual Risk Analyser ---
    contracts_clause = ""
    if enable_contracts:
        contracts_clause = textwrap.dedent("""
            CONTRACTUAL RISK ANALYSIS:
            Scan all documents for contract language (DPA, MSA, SLA, NDA, service agreements).
            Check each against this golden-clause checklist and flag missing or inadequate clauses:
              1. Right to audit / right to inspect (at least annually, on reasonable notice)
              2. Breach/incident notification SLA (must be ≤72h for GDPR Article 33 compliance)
              3. Sub-processor restriction and prior written approval requirement
              4. Data deletion / return obligation on termination (including backups)
              5. Liability cap adequacy (must cover at minimum 12 months of contract value)
              6. BCP/DRP obligation and annual testing requirement
              7. Regulatory compliance warranty (GDPR, NIS2, DORA as applicable)
              8. Intellectual property and data ownership clause
              9. Security standards commitment (ISO 27001 / SOC 2 certification or equivalent)
             10. Governing law and jurisdiction clarity
            For each absent or weak clause, describe the legal/operational risk it creates.
        """).strip()

    # --- Feature B: Red Flag Detector ---
    red_flag_clause = ""
    if enable_red_flag:
        red_flag_clause = textwrap.dedent("""
            RED FLAG DETECTION:
            Perform a fast pass specifically looking for deal-breaker signals:
              - Evidence or admission of a past data breach or security incident
              - Named unpatched CVEs or end-of-life software in production
              - Regulatory sanctions, fines, or active enforcement actions
              - Complete absence of fundamental controls (no encryption, no access control, no logging)
              - Major inconsistencies between stated policies and observed technical setup
              - Single points of failure in critical infrastructure with no documented DR
            List only genuine, evidence-based red flags. If none found, state "No critical red flags identified."
        """).strip()

    # --- Feature C: Supply Chain Mapper ---
    supply_chain_clause = ""
    if enable_supply_chain:
        supply_chain_clause = textwrap.dedent("""
            SUPPLY CHAIN MAPPING:
            Extend the ThirdParties section into a full multi-tier supply chain map.
            For each identified third party and sub-processor:
              - Classify by tier: Tier 1 (direct), Tier 2 (sub-processor of sub-processor)
              - Identify service category: Cloud/Hosting, Identity, Monitoring, Payments, Communications, CDN, Other
              - Flag concentration risks: geographic concentration, single-vendor dependency, hyperscaler lock-in
              - Estimate data access level: None / Indirect / Direct access to personal or sensitive data
            Produce an overall supply chain risk rating (Low/Medium/High/Critical) with justification.
        """).strip()

    # --- Feature D: Follow-up Questionnaire ---
    questionnaire_clause = ""
    if enable_questionnaire:
        questionnaire_clause = textwrap.dedent("""
            FOLLOW-UP QUESTIONNAIRE:
            Based solely on the gaps and risks identified, generate a targeted follow-up questionnaire.
            For each question:
              - Reference the specific gap or control failure it addresses
              - State the ISO 27001 Annex A or NIST CSF control it maps to
              - Specify the expected evidence type (policy document, audit report, screenshot, certificate)
            Generate between 10 and 20 questions. Do not include generic or boilerplate questions.
            Only ask what is genuinely missing or unclear from the provided documents.
        """).strip()

    # --- Feature E: Data Sensitivity Heatmap ---
    heatmap_clause = ""
    if enable_heatmap:
        heatmap_clause = textwrap.dedent("""
            DATA SENSITIVITY HEATMAP:
            Classify all data types the supplier processes based on evidence in the documents.
            For each identified data type:
              - Assign a sensitivity tier: Public / Internal / Confidential / Restricted / Special Category
              - Note the applicable regulation: GDPR, PCI-DSS, HIPAA, NIS2, DORA, or N/A
              - Flag if cross-border transfers are involved and identify destination countries
              - Indicate the retention period if documented
            Conclude with a DPIA trigger assessment: state whether a formal Data Protection Impact
            Assessment or Legitimate Interest Assessment is required, and why.
        """).strip()

    # --- A–E schema fragments ---
    contracts_schema = ""
    if enable_contracts:
        contracts_schema = textwrap.dedent("""
            ,
              "ContractualRiskAnalysis": {
                "DocumentsReviewed": "string – list of contractual documents found and reviewed",
                "MissingClauses": [
                  {
                    "ClauseTitle": "string",
                    "GoldenStandard": "string – what the clause should say",
                    "CurrentStatus": "Absent / Insufficient / Adequate",
                    "RiskIfAbsent": "string – legal or operational exposure"
                  }
                ],
                "ContractualRiskRating": "Low / Medium / High / Critical",
                "ContractualSummary": "string – overall narrative on contractual posture"
              }
        """).strip()

    supply_chain_schema = ""
    if enable_supply_chain:
        supply_chain_schema = textwrap.dedent("""
            ,
              "SupplyChainMap": {
                "Parties": [
                  {
                    "Name": "string",
                    "Tier": "Tier 1 / Tier 2",
                    "ServiceCategory": "string",
                    "DataAccessLevel": "None / Indirect / Direct",
                    "GeographicLocation": "string",
                    "ConcentrationRisk": "string – specific risk or 'None identified'"
                  }
                ],
                "OverallSupplyChainRisk": "Low / Medium / High / Critical",
                "SupplyChainSummary": "string – narrative on key concentration risks and dependencies"
              }
        """).strip()

    heatmap_schema = ""
    if enable_heatmap:
        heatmap_schema = textwrap.dedent("""
            ,
              "DataSensitivityHeatmap": {
                "DataTypes": [
                  {
                    "DataType": "string – e.g. Employee PII, Payment Card Data",
                    "SensitivityTier": "Public / Internal / Confidential / Restricted / Special Category",
                    "Regulation": "string – GDPR / PCI-DSS / HIPAA / NIS2 / N/A",
                    "CrossBorderTransfer": "boolean",
                    "TransferDestinations": "string",
                    "RetentionPeriod": "string"
                  }
                ],
                "DPIARequired": "boolean",
                "DPIARationale": "string",
                "HeatmapSummary": "string"
              }
        """).strip()

    # Red flag and questionnaire go into their own top-level keys (always included if enabled)
    red_flag_schema = ""
    if enable_red_flag:
        red_flag_schema = textwrap.dedent("""
            ,
              "RedFlags": [
                {
                  "Flag": "string – concise one-line description",
                  "Severity": "High / Critical",
                  "Evidence": "string – document name + specific finding",
                  "Implication": "string – what this means for the engagement"
                }
              ]
        """).strip()

    questionnaire_schema = ""
    if enable_questionnaire:
        questionnaire_schema = textwrap.dedent("""
            ,
              "FollowUpQuestionnaire": [
                {
                  "Question": "string",
                  "GapAddressed": "string – which risk or missing control this targets",
                  "ControlReference": "string – e.g. ISO 27001 A.9.4 / NIST CSF PR.AC-1",
                  "ExpectedEvidence": "string – e.g. Policy document, SOC 2 report, Screenshot"
                }
              ]
        """).strip()

    # Re-assemble full schema with ALL feature fragments
    json_schema = (
        "OUTPUT JSON SCHEMA (return ONLY this JSON, nothing else):\n"
        + core_schema
        + effort_schema
        + stack_schema
        + scorecard_schema
        + contracts_schema
        + red_flag_schema
        + supply_chain_schema
        + questionnaire_schema
        + heatmap_schema
        + "\n}"
    )

    # ------------------------------------------------------------------
    # Build instructions
    # ------------------------------------------------------------------
    instruction_parts = [
        "INSTRUCTIONS:",
        "1. First scan all documents for security questionnaires or compliance checklists."
        "\n   Map every control that is marked 'No', 'Not implemented', 'Partial', or equivalent."
        "\n   These gaps form the factual backbone of the risk register.",
        "2. Cross-reference with the rest of the documents to identify additional risks,"
        "\n   inconsistencies, and evidence of past incidents.",
        "3. Be specific: name the exact technology, configuration, or process that creates each risk."
        "\n   Never write generic statements like 'the company should improve its security posture.'",
        "4. For each risk in RisksAndRecommendations, the Sources field must reference ONLY"
        "\n   information found in the provided documents (file name + section/page).",
        "5. The RemediationPlan must be derived directly from the identified risks and prioritised"
        "\n   by risk level (Critical → High → Medium → Low).",
        "6. Provide continuous prose (not sub-lists) for the narrative sections"
        "\n   (Services, TechnicalSetup, SecurityMeasures, DataFlows, ThirdParties,"
        "\n   GovernanceRecommendations, OwnershipMatrix, IntegrationSummary,"
        "\n   ContractualSummary, SupplyChainSummary, HeatmapSummary).",
        "7. Output only valid JSON matching the schema above.",
    ]
    if enable_effort:
        instruction_parts.append(
            "8. EffortEstimate: cover every action in RemediationPlan. Apply org-size factor "
            f"{ORG_SIZE_PROFILES.get(org_size, ORG_SIZE_PROFILES['Mid'])['effort_factor']}× to mid estimates."
        )
    if acquirer_stack:
        instruction_parts.append(
            "9. ITStackComparison: extract supplier stack from documents first. "
            "Mark undocumented tools as 'Not documented'."
        )
    if enable_scorecard:
        instruction_parts.append(
            "10. ExecutiveScorecard: cite at least one source per domain. "
            "OverallRating = Red if any High/Critical-risk domain is Red."
        )
    if enable_contracts:
        instruction_parts.append(
            "11. ContractualRiskAnalysis: check EVERY clause in the golden list. "
            "Mark as Absent only if genuinely not found."
        )
    if enable_red_flag:
        instruction_parts.append(
            "12. RedFlags: list only evidence-based deal-breakers. "
            "If none, return an empty array — do NOT fabricate flags."
        )
    if enable_supply_chain:
        instruction_parts.append(
            "13. SupplyChainMap: include every named third party. "
            "Flag hyperscaler concentration where both parties use the same provider."
        )
    if enable_questionnaire:
        instruction_parts.append(
            "14. FollowUpQuestionnaire: 10–20 targeted questions only. "
            "Each must map to a specific gap found in the documents."
        )
    if enable_heatmap:
        instruction_parts.append(
            "15. DataSensitivityHeatmap: classify every data type mentioned. "
            "Set DPIARequired=true if special category data or large-scale processing is evident."
        )

    instruction = "\n".join(instruction_parts)

    documents_block = f"DOCUMENTS TO ANALYSE:\n\n{combined_text}"

    return "\n\n".join(
        part for part in [
            context_clause.strip(),
            reg_clause.strip(),
            effort_clause,
            stack_clause,
            scorecard_clause,
            contracts_clause,
            red_flag_clause,
            supply_chain_clause,
            questionnaire_clause,
            heatmap_clause,
            ebios_scale,
            json_schema,
            instruction,
            documents_block,
        ]
        if part
    )


# ---------------------------------------------------------------------------
# JSON extraction helper
# ---------------------------------------------------------------------------

def extract_json(raw: str) -> dict:
    """
    Robustly extract a JSON object from LLM output that may contain
    markdown fences or surrounding text.
    """
    # Try direct parse first
    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        pass

    # Strip markdown fences
    stripped = raw
    for fence in ("```json", "```"):
        if fence in stripped:
            stripped = stripped.split(fence, 1)[-1]
            stripped = stripped.rsplit("```", 1)[0]
            break

    try:
        return json.loads(stripped.strip())
    except json.JSONDecodeError:
        pass

    # Find the outermost JSON object
    start = raw.find("{")
    end = raw.rfind("}")
    if start != -1 and end != -1:
        try:
            return json.loads(raw[start : end + 1])
        except json.JSONDecodeError:
            pass

    # Give up gracefully
    print("[WARN] Could not parse LLM output as JSON. Returning raw text under 'RawOutput'.")
    return {"RawOutput": raw}


# ---------------------------------------------------------------------------
# DOCX memorandum generation
# ---------------------------------------------------------------------------

def _fmt_list(items: list[dict], fields: list[str]) -> str:
    """Format a list of dicts into readable prose for the DOCX template."""
    lines = []
    for i, item in enumerate(items, 1):
        parts = [f"{i}."]
        for field in fields:
            val = item.get(field, "")
            if val:
                parts.append(f"[{field}] {val}")
        lines.append("  ".join(parts))
    return "\n\n".join(lines)


def create_memorandum(analysis: dict, template_path: str, output_path: str) -> None:
    doc = DocxTemplate(template_path)

    risks = analysis.get("RisksAndRecommendations", [])
    risks_and_recommendations = [
        {
            "title": r.get("RiskTitle", "N/A"),
            "description": r.get("RiskDescription", "N/A"),
            "likelihood": r.get("Likelihood", "N/A"),
            "impact": r.get("Impact", "N/A"),
            "risk_level": r.get("RiskLevel", "N/A"),
            "recommendations": r.get("Recommendations", "N/A"),
            "sources": r.get("Sources", "N/A"),
        }
        for r in risks
    ]

    remediation = analysis.get("RemediationPlan", {})

    priority_actions = [
        {
            "action": a.get("Action", "N/A"),
            "rationale": a.get("Rationale", "N/A"),
            "owner": a.get("Owner", "N/A"),
            "deadline": a.get("Deadline", "0–30 days"),
        }
        for a in remediation.get("PriorityActions", [])
    ]

    medium_term_actions = [
        {
            "action": a.get("Action", "N/A"),
            "rationale": a.get("Rationale", "N/A"),
            "owner": a.get("Owner", "N/A"),
            "deadline": a.get("Deadline", "1–6 months"),
        }
        for a in remediation.get("MediumTermActions", [])
    ]

    strategic_actions = [
        {
            "action": a.get("Action", "N/A"),
            "rationale": a.get("Rationale", "N/A"),
            "owner": a.get("Owner", "N/A"),
            "deadline": a.get("Deadline", "6–18 months"),
        }
        for a in remediation.get("StrategicActions", [])
    ]

    # --- Feature 2: Effort Estimator ---
    effort = analysis.get("EffortEstimate", {})
    effort_action_estimates = [
        {
            "action": a.get("Action", "N/A"),
            "phase": a.get("Phase", "N/A"),
            "effort_days_min": a.get("EffortDaysMin", "N/A"),
            "effort_days_mid": a.get("EffortDaysMid", "N/A"),
            "effort_days_max": a.get("EffortDaysMax", "N/A"),
            "role_profile": a.get("RoleProfile", "N/A"),
            "in_house_feasible": "Yes" if a.get("InHouseFeasible") else "External support recommended",
            "notes": a.get("Notes", ""),
        }
        for a in effort.get("ActionEstimates", [])
    ]

    # --- Feature 1: IT Stack Comparator ---
    stack_cmp = analysis.get("ITStackComparison", {})
    supplier_stack = stack_cmp.get("SupplierStack", {})
    compatibility_matrix = [
        {
            "layer": row.get("Layer", "N/A"),
            "supplier_tool": row.get("SupplierTool", "N/A"),
            "acquirer_tool": row.get("AcquirerTool", "N/A"),
            "status": row.get("CompatibilityStatus", "N/A"),
            "hurdle": row.get("IntegrationHurdle", ""),
            "opportunity": row.get("ConsolidationOpportunity", ""),
        }
        for row in stack_cmp.get("CompatibilityMatrix", [])
    ]

    # --- Feature 3: Executive Scorecard ---
    scorecard = analysis.get("ExecutiveScorecard", {})
    scorecard_domains = [
        {
            "domain": d.get("Domain", "N/A"),
            "score": d.get("Score", "N/A"),
            "delta": d.get("Delta", "N/A"),
            "key_finding": d.get("KeyFinding", "N/A"),
            "evidence_source": d.get("EvidenceSource", "N/A"),
        }
        for d in scorecard.get("Domains", [])
    ]

    context = {
        # Header fields
        "title": "Cybersecurity Risk Assessment Memorandum",
        "company_name": analysis.get("CompanyName", ""),
        "assessment_date": analysis.get("AssessmentDate", ""),
        "introduction": (
            "This memorandum provides a comprehensive cybersecurity risk assessment based on the "
            "analysis of the provided documentation. It covers the supplier's services, technical "
            "architecture, security controls, data flows, sub-processors, identified risks, and a "
            "post-assessment remediation roadmap."
        ),
        # Assessment sections
        "services": analysis.get("Services", "No data available."),
        "technical_setup": analysis.get("TechnicalSetup", "No data available."),
        "security_measures": analysis.get("SecurityMeasures", "No data available."),
        "data_flows": analysis.get("DataFlows", "No data available."),
        "third_parties": analysis.get("ThirdParties", "No data available."),
        # Risk register
        "risks_and_recommendations": risks_and_recommendations,
        # Remediation plan
        "priority_actions": priority_actions,
        "medium_term_actions": medium_term_actions,
        "strategic_actions": strategic_actions,
        "governance_recommendations": remediation.get(
            "GovernanceRecommendations", "No data available."
        ),
        "ownership_matrix": remediation.get("OwnershipMatrix", "No data available."),
        # Feature 2 – Effort Estimator
        "effort_org_size": effort.get("OrgSize", ""),
        "effort_total_min": effort.get("TotalEffortDaysMin", ""),
        "effort_total_mid": effort.get("TotalEffortDaysMid", ""),
        "effort_total_max": effort.get("TotalEffortDaysMax", ""),
        "effort_key_assumptions": effort.get("KeyAssumptions", ""),
        "effort_action_estimates": effort_action_estimates,
        # Feature 1 – IT Stack Comparator
        "supplier_stack_iam": supplier_stack.get("IAM", ""),
        "supplier_stack_siem": supplier_stack.get("SIEM", ""),
        "supplier_stack_itsm": supplier_stack.get("ITSM", ""),
        "supplier_stack_cloud": supplier_stack.get("Cloud", ""),
        "supplier_stack_network": supplier_stack.get("Network", ""),
        "supplier_stack_endpoint": supplier_stack.get("Endpoint", ""),
        "supplier_stack_data": supplier_stack.get("DataPlatform", ""),
        "supplier_stack_bcdr": supplier_stack.get("BCDR", ""),
        "compatibility_matrix": compatibility_matrix,
        "integration_risk": stack_cmp.get("OverallIntegrationRisk", ""),
        "integration_summary": stack_cmp.get("IntegrationSummary", ""),
        # Feature 3 – Executive Scorecard
        "scorecard_overall_rating": scorecard.get("OverallRating", ""),
        "scorecard_overall_summary": scorecard.get("OverallSummary", ""),
        "scorecard_domains": scorecard_domains,
        # Feature A – Contractual Risk Analyser
        "contractual_docs_reviewed": analysis.get("ContractualRiskAnalysis", {}).get("DocumentsReviewed", ""),
        "contractual_risk_rating": analysis.get("ContractualRiskAnalysis", {}).get("ContractualRiskRating", ""),
        "contractual_summary": analysis.get("ContractualRiskAnalysis", {}).get("ContractualSummary", ""),
        "missing_clauses": [
            {
                "clause_title": c.get("ClauseTitle", "N/A"),
                "golden_standard": c.get("GoldenStandard", "N/A"),
                "current_status": c.get("CurrentStatus", "N/A"),
                "risk_if_absent": c.get("RiskIfAbsent", "N/A"),
            }
            for c in analysis.get("ContractualRiskAnalysis", {}).get("MissingClauses", [])
        ],
        # Feature B – Red Flag Detector
        "red_flags": [
            {
                "flag": f.get("Flag", "N/A"),
                "severity": f.get("Severity", "N/A"),
                "evidence": f.get("Evidence", "N/A"),
                "implication": f.get("Implication", "N/A"),
            }
            for f in analysis.get("RedFlags", [])
        ],
        # Feature C – Supply Chain Mapper
        "supply_chain_risk": analysis.get("SupplyChainMap", {}).get("OverallSupplyChainRisk", ""),
        "supply_chain_summary": analysis.get("SupplyChainMap", {}).get("SupplyChainSummary", ""),
        "supply_chain_parties": [
            {
                "name": p.get("Name", "N/A"),
                "tier": p.get("Tier", "N/A"),
                "service_category": p.get("ServiceCategory", "N/A"),
                "data_access": p.get("DataAccessLevel", "N/A"),
                "location": p.get("GeographicLocation", "N/A"),
                "concentration_risk": p.get("ConcentrationRisk", ""),
            }
            for p in analysis.get("SupplyChainMap", {}).get("Parties", [])
        ],
        # Feature D – Follow-up Questionnaire
        "followup_questions": [
            {
                "question": q.get("Question", "N/A"),
                "gap_addressed": q.get("GapAddressed", "N/A"),
                "control_reference": q.get("ControlReference", "N/A"),
                "expected_evidence": q.get("ExpectedEvidence", "N/A"),
            }
            for q in analysis.get("FollowUpQuestionnaire", [])
        ],
        # Feature E – Data Sensitivity Heatmap
        "dpia_required": "Yes" if analysis.get("DataSensitivityHeatmap", {}).get("DPIARequired") else "No",
        "dpia_rationale": analysis.get("DataSensitivityHeatmap", {}).get("DPIARationale", ""),
        "heatmap_summary": analysis.get("DataSensitivityHeatmap", {}).get("HeatmapSummary", ""),
        "data_types": [
            {
                "data_type": d.get("DataType", "N/A"),
                "sensitivity_tier": d.get("SensitivityTier", "N/A"),
                "regulation": d.get("Regulation", "N/A"),
                "cross_border": "Yes" if d.get("CrossBorderTransfer") else "No",
                "destinations": d.get("TransferDestinations", ""),
                "retention": d.get("RetentionPeriod", ""),
            }
            for d in analysis.get("DataSensitivityHeatmap", {}).get("DataTypes", [])
        ],
    }

    doc.render(context)
    doc.save(output_path)
    print(f"\n[OK] Memorandum saved to: {output_path}")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main() -> None:
    parser = argparse.ArgumentParser(
        description="TPSA – LLM-agnostic Third-Party Security Assessment Generator",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=textwrap.dedent("""
            Examples:
              python main.py ./docs ./template.docx ./report.docx
              python main.py ./docs ./template.docx ./report.docx --provider anthropic --model claude-opus-4-5
              python main.py ./docs ./template.docx ./report.docx --provider gemini --context ma --regulations nis2,dora
              python main.py ./docs ./template.docx ./report.docx --effort --org-size Mid
              python main.py ./docs ./template.docx ./report.docx --stack ./acquirer_stack.json
              python main.py ./docs ./template.docx ./report.docx --scorecard --baseline ./previous_report.json
        """),
    )

    # Positional arguments (backwards compatible)
    parser.add_argument("folder_path", help="Folder containing PDF/DOCX documents to analyse")
    parser.add_argument("template_path", help="DOCX template file path")
    parser.add_argument("output_path", help="Output DOCX file path")

    # LLM configuration
    parser.add_argument(
        "--provider", "-p",
        choices=[p.value for p in Provider],
        default=Provider.OPENAI.value,
        help="LLM provider to use (default: openai)",
    )
    parser.add_argument(
        "--model", "-m",
        default=None,
        help="Model name (default per provider: gpt-4o / claude-opus-4-5 / gemini-1.5-pro / mistral-large-latest)",
    )
    parser.add_argument(
        "--api-key", "-k",
        default=None,
        dest="api_key",
        help="API key. Falls back to OPENAI_API_KEY / ANTHROPIC_API_KEY / GEMINI_API_KEY / MISTRAL_API_KEY env vars",
    )

    # Assessment scope
    parser.add_argument(
        "--context", "-c",
        choices=["tprm", "ma"],
        default="tprm",
        help="Assessment context: 'tprm' (default) or 'ma' (M&A due diligence)",
    )
    parser.add_argument(
        "--regulations", "-r",
        default="",
        help="Comma-separated regulatory frameworks to assess against, e.g.: nis2,dora,soc2,iso27001,gdpr",
    )
    parser.add_argument(
        "--save-raw",
        action="store_true",
        help="Save raw LLM JSON output alongside the DOCX (adds .json extension)",
    )

    # Feature 2: Effort Estimator
    effort_group = parser.add_argument_group("Feature: Effort Estimator")
    effort_group.add_argument(
        "--effort",
        action="store_true",
        help="Enable remediation effort estimation (man-days per action with role profiles)",
    )
    effort_group.add_argument(
        "--org-size",
        choices=ORG_SIZE_LABELS,
        default="Mid",
        dest="org_size",
        help="Organisation size for effort calibration: SME / Mid (default) / Large / Enterprise",
    )

    # Feature 1: IT Stack Comparator
    stack_group = parser.add_argument_group("Feature: IT Stack Comparator")
    stack_group.add_argument(
        "--stack",
        default=None,
        metavar="STACK_JSON",
        help=(
            "Path to a JSON file describing the acquirer's IT/security stack. "
            "Triggers the IT stack comparison feature. "
            "Expected keys: IAM, SIEM, ITSM, Cloud, Network, Endpoint, DataPlatform, BCDR."
        ),
    )

    # Feature 3: Executive Scorecard
    scorecard_group = parser.add_argument_group("Feature: Executive Scorecard")
    scorecard_group.add_argument(
        "--scorecard",
        action="store_true",
        help="Generate a RAG executive scorecard across eight security domains",
    )
    scorecard_group.add_argument(
        "--baseline",
        default=None,
        metavar="BASELINE_JSON",
        help=(
            "Path to a previous assessment JSON (--save-raw output) for delta comparison. "
            "Requires --scorecard."
        ),
    )

    # Feature A: Contractual Risk Analyser
    contracts_group = parser.add_argument_group("Feature: Contractual Risk Analyser")
    contracts_group.add_argument(
        "--contracts",
        action="store_true",
        help="Analyse contracts/DPAs/SLAs against a 10-point golden-clause checklist",
    )

    # Feature B: Red Flag Detector
    redflag_group = parser.add_argument_group("Feature: Red Flag Detector")
    redflag_group.add_argument(
        "--red-flags",
        action="store_true",
        dest="red_flags",
        help="Run a fast pass to identify critical deal-breaker signals",
    )

    # Feature C: Supply Chain Mapper
    supplychain_group = parser.add_argument_group("Feature: Supply Chain Mapper")
    supplychain_group.add_argument(
        "--supply-chain",
        action="store_true",
        dest="supply_chain",
        help="Generate a multi-tier supply chain map with concentration risk analysis",
    )

    # Feature D: Follow-up Questionnaire
    questionnaire_group = parser.add_argument_group("Feature: Follow-up Questionnaire")
    questionnaire_group.add_argument(
        "--questionnaire",
        action="store_true",
        help="Generate a targeted follow-up questionnaire based on identified gaps",
    )

    # Feature E: Data Sensitivity Heatmap
    heatmap_group = parser.add_argument_group("Feature: Data Sensitivity Heatmap")
    heatmap_group.add_argument(
        "--heatmap",
        action="store_true",
        help="Classify data types by sensitivity tier and assess DPIA trigger",
    )

    args = parser.parse_args()

    provider = Provider(args.provider)
    model = args.model or DEFAULT_MODELS[provider]
    api_key = _resolve_api_key(provider, args.api_key)
    regulations = [r.strip().lower() for r in args.regulations.split(",") if r.strip()]

    # Load optional JSON inputs
    acquirer_stack: Optional[dict] = None
    if args.stack:
        stack_path = Path(args.stack)
        if not stack_path.is_file():
            sys.exit(f"[ERROR] Stack file not found: {args.stack}")
        with open(stack_path, encoding="utf-8") as f:
            acquirer_stack = json.load(f)
        print(f"[INFO] Acquirer stack loaded from: {args.stack}")

    baseline: Optional[dict] = None
    if args.baseline:
        if not args.scorecard:
            print("[WARN] --baseline has no effect without --scorecard.")
        else:
            baseline_path = Path(args.baseline)
            if not baseline_path.is_file():
                sys.exit(f"[ERROR] Baseline file not found: {args.baseline}")
            with open(baseline_path, encoding="utf-8") as f:
                baseline = json.load(f)
            print(f"[INFO] Baseline assessment loaded from: {args.baseline}")

    # Print run summary
    active_features = []
    if args.effort:
        active_features.append(f"Effort Estimator [{args.org_size}]")
    if acquirer_stack:
        active_features.append("IT Stack Comparator")
    if args.scorecard:
        active_features.append("Executive Scorecard" + (" [delta]" if baseline else ""))
    if args.contracts:
        active_features.append("Contractual Risk Analyser")
    if args.red_flags:
        active_features.append("Red Flag Detector")
    if args.supply_chain:
        active_features.append("Supply Chain Mapper")
    if args.questionnaire:
        active_features.append("Follow-up Questionnaire")
    if args.heatmap:
        active_features.append("Data Sensitivity Heatmap")

    print("\n" + "=" * 60)
    print("  TPSA – Third-Party Security Assessment Generator")
    print("=" * 60)
    print(f"  Provider   : {provider.value}")
    print(f"  Model      : {model}")
    print(f"  Context    : {args.context.upper()}")
    if regulations:
        print(f"  Regulations: {', '.join(regulations).upper()}")
    if active_features:
        print(f"  Features   : {', '.join(active_features)}")
    print(f"  Documents  : {args.folder_path}")
    print(f"  Output     : {args.output_path}")
    print("=" * 60)

    # Step 1 – Extract document text
    combined_text = collect_text_from_folder(args.folder_path)
    print(f"\n[INFO] Total extracted text: {len(combined_text):,} characters")

    # Step 2 – Build prompt (active features injected into schema dynamically)
    user_prompt = build_user_prompt(
        combined_text=combined_text,
        context_mode=args.context,
        regulations=regulations,
        enable_effort=args.effort,
        org_size=args.org_size,
        acquirer_stack=acquirer_stack,
        enable_scorecard=args.scorecard,
        baseline=baseline,
        enable_contracts=args.contracts,
        enable_red_flag=args.red_flags,
        enable_supply_chain=args.supply_chain,
        enable_questionnaire=args.questionnaire,
        enable_heatmap=args.heatmap,
    )

    # Step 3 – Call LLM
    raw_output = call_llm(SYSTEM_PROMPT, user_prompt, provider, model, api_key)

    # Step 4 – Parse JSON
    analysis = extract_json(raw_output)

    # Optionally save raw JSON
    if args.save_raw:
        raw_path = args.output_path.replace(".docx", ".json")
        with open(raw_path, "w", encoding="utf-8") as f:
            json.dump(analysis, f, indent=2, ensure_ascii=False)
        print(f"[INFO] Raw JSON saved to: {raw_path}")

    # Step 5 – Generate DOCX
    create_memorandum(analysis, args.template_path, args.output_path)


if __name__ == "__main__":
    main()

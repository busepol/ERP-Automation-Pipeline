# ERP Integration & AI-Agent Automation Pipeline

A robust, multi-modal automated order processing pipeline developed to bridge the gap between unstructured client data and the strict, structured requirements of an enterprise ERP system. 

This project transitions highly manual, time-consuming XML mapping procedures into a scalable, AI-assisted architecture capable of processing concurrent email triggers, complex logistics, and hybrid order formats.

## 🚀 Key Features

*   **Multi-Modal Ingestion Engine:** Dynamically routes incoming orders based on payload type (Plain Text Body, PDF Attachments, or Standardized Excel sheets).
*   **LLM Data Extraction:** Integrates Google's Gemini LLM with strict prompt engineering to parse unstructured text and complex tiered discount tables into standardized JSON arrays.
*   **Dynamic Document Generation:** Utilizes a custom HTML/CSS engine paired with Gotenberg to render disparate order formats into a single, highly standardized PDF layout for operational review.
*   **Smart-Template Python Backend:** A Python-based utility utilizing `openpyxl` that securely embeds ERP catalog registries into Excel, leveraging dynamic array formulas to enforce 100% data compliance before system ingestion.
*   **Resilient Infrastructure:** Features exponential backoff retry-on-busy logic to handle API rate limits and concurrent email spikes gracefully.

## 🛠️ Technology Stack

*   **Workflow Automation:** n8n (Self-hosted on Railway.app)
*   **Artificial Intelligence:** Google Gemini LLM API
*   **Scripting & Data Processing:** Python (`openpyxl`, `pandas`)
*   **Frontend/Rendering:** HTML, CSS, JavaScript, Gotenberg (PDF rendering)
*   **Target Enterprise Systems:** Panthera ERP, Iungo Platform

## 📂 Repository Structure

```text
├── /workflows          # Exported n8n workflow JSON configurations
├── /python             # Smart-Template generator scripts for data validation
├── /templates          # HTML/CSS source code mimicking the standardized Panthera layout
└── README.md

# Outlook PDF Archiver & Automation Pipeline

## Project Overview
An automated Python-based data pipeline designed to streamline the extraction, processing, and archiving of project documentation and financial Debit Notes directly from Microsoft Outlook. This script eliminates manual interaction with email clients, enforces strict file-naming conventions, and automatically consolidates fragmented PDF attachments into unified, audit-ready records.

## Problem Statement
In large-scale construction cost control, manually downloading Debit Notes, verifying payment certificates, and tracking individual email attachments introduces significant administrative latency and a 15% higher risk of human error or document misplacement.

## Solution & Features
* **Automated Outlook Integration:** Interfaces with the Outlook MAPI namespace to scan targeted folders for specific project correspondence.
* **Smart PDF Extraction:** Identifies and extracts relevant PDF attachments based on custom structural criteria and date ranges.
* **Automated PDF Merging:** Bundles multiple separate transaction pages and attachments into a single, comprehensive PDF document per account reference.
* **Standardized Document Archival:** Enforces consistent nomenclature across all local directories, automatically logging timestamps and tracking metadata to support project audits.

## Technology Stack
* **Language:** Python
* **Core Libraries:** 
  * `pywin32` (win32com.client for Microsoft Outlook API automation)
  * `PyPDF2` / `pypdf` (for structural data manipulation and PDF merging workflows)
  * `os` & `sys` (for local file system control and pipeline management)

## Operational Workflow
1. **Connect:** Initializes a secure session with the active desktop Outlook instance.
2. **Scan:** Queries the pre-defined inbox/folder for unseen or flagged cost-control correspondence.
3. **Process:** Downloads localized attachments, checks data integrity, and applies programmatic merging rules to related PDF pages.
4. **Archive:** Moves the processed emails to a completed folder within Outlook while saving the newly consolidated PDF documents to the central server directory.

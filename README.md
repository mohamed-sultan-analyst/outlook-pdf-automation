# Outlook PDF Archiver & Debit Note Automation Pipeline

## Project Overview
An advanced Python automation tool designed for Cost Control operations to streamline the verification, archiving, and dashboard integration of Debit Notes directly from Microsoft Outlook. The pipeline automates the generation of approval cover sheets via programmatic screenshots, dynamically prepends them to their corresponding financial PDFs, and routes the finalized documents to an integrated central directory synchronized with Microsoft Excel dashboards for single-click cross-verification.

## Problem Statement
In mega-project cost control, manually reconciling approved emails with Debit Note PDFs and updating tracking ledgers introduces severe processing delays. Without systemic synchronization, senior management lacks immediate, auditable access to validation cover sheets directly from financial summary reports, creating data silos and slowing down the audit workflow.

## Solution & Features
* **Programmatic Approval Capturing:** Automatically scans Outlook for approved Debit Note emails and generates programmatic screenshots to serve as the official document cover page.
* **Smart Document Matching:** Searches and identifies the local PDF file matching the exact name and reference number found in the verified email.
* **Dynamic PDF Prepending:** Automatically inserts the email approval screenshot as the first page (Cover) of the target Debit Note PDF, ensuring audit-ready documentation.
* **Excel Dashboard Synchronization:** Routes the fully compiled PDF directly into the destination directory linked with Excel tracking systems, enabling stakeholders to instantly open and verify the approved PDF with a single click from the dashboard.
* **Terminal Output Logging:** Prints real-time tracking metrics directly to the execution terminal, displaying the exact email profiles and extracted Debit Note reference numbers being processed.

## Technology Stack
* **Language:** Python
* **Core Libraries:** 
  * `pywin32` (win32com.client for Microsoft Outlook API automation)
  * `pypdf` / `PyPDF2` (for dynamic page insertion and PDF manipulation)
  * `Pillow` / `pyautogui` (for programmatic approval screenshot generation)
  * `os` & `sys` (for system directory and local server path management)

## Operational Workflow
1. **Fetch & Scan:** Connects to the active Outlook instance to scan verified correspondence for Debit Notes and their unique reference numbers.
2. **Capture Cover:** Takes an automated screenshot of the approved email to establish the official cover page.
3. **Match & Prepend:** Locates the matching local Debit Note PDF by name and inserts the captured screenshot as the first page.
4. **Data Integration:** Saves the completed record directly into the source folder queried by the Excel monitoring dashboard to enable dynamic hyperlink verification.
5. **Terminal Review:** Displays active processing status and compiled figures directly on the terminal screen for immediate verification.

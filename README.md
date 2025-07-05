# Consultation Crusher

This tool automates extraction, analysis, and completion of consultation-related tasks from an internal system. It uses Playwright for browser automation to log into the system, scrape task data, summarize job types, and finalize tasks with detailed dispatch notes.

---

## Overview

- **Task Extraction**: Automatically loads due consultation tasks from a configured internal URL.
- **Job Type Classification**: Uses fuzzy string matching to classify jobs as Free, Billable, or Unknown based on job descriptions.
- **Dispatch Summary Generation**: Visits linked work orders and customer pages to generate detailed dispatch summaries, including arrival/departure times, equipment used, and responsible party.
- **Task Finalization**: Fills task notes with the generated summary, marks tasks completed, and optionally spawns billing subtasks.
- **Session Management**: Supports persistent login sessions with stored state files and automatic Playwright Chromium installation.
- **Logging & Progress Bar**: Logs detailed progress and errors with timestamps, and provides a progress bar for task processing.
- **CLI Options**: Supports `--version` flag for version info.

---

## Features

- **Headless Browser Automation** using Playwright (Chromium) for interacting with the internal site.
- **Credential Handling** via environment variables or interactive prompts, with secure password entry.
- **Dynamic Popup Handling** to dismiss first-time overlays or modal dialogs encountered on login or navigation.
- **Fuzzy Matching** of job types using RapidFuzz for robust classification.
- **Detailed Notes Extraction** and HTML sanitization to generate clean, readable dispatch summaries.
- **Error Handling** and debugging helpers to diagnose issues with page navigation or missing elements.
- **Configurable Timeouts** and logging to handle slow network or UI delays gracefully.

---

## Typical Workflow

1. **Setup & Launch**  
   - Install dependencies from `requirements.txt`.  
   - Run the script `backend.py` directly or via bundled executable.

2. **Login**  
   - The tool attempts to restore session from saved state.  
   - If no session exists, prompts for username and password securely.  
   - Saves session state for future runs.

3. **Extract Due Consultation Tasks**  
   - Navigates to the configured task URL.  
   - Scrapes tasks due today or earlier, filtering for those with "consultation" in the description.

4. **Process Each Task**  
   - Parses job type from task notes using fuzzy matching and regex patterns.  
   - Generates dispatch summaries by visiting customer and work order pages.  
   - Expands task forms and fills notes with the summary.  
   - Marks task as completed, optionally spawning billing subtasks for billable dispatches.

5. **Summary & Logging**  
   - Logs task processing results with timestamps.  
   - Prints a summary count of processed job types.  
   - Saves detailed logs to a timestamped log file.

---

## Configuration & Environment

- **Environment Variables** stored in `.env` file under the `Misc` directory, including:  
  - `UNITY_USER` and `PASSWORD` for credentials.  
- **Playwright Browser Binaries** are downloaded automatically on first run into a configurable browsers folder.  
- **Logging** writes to `logs/consulation_log.txt` with timestamps and detailed messages.

---

## Requirements

- Python 3.10+  
- Packages listed in `requirements.txt`, including but not limited to:  
  - playwright  
  - python-dotenv  
  - rapidfuzz  
  - tqdm  
  - pandas, numpy, openpyxl (optional, for any data processing)  

---

## Limitations & Notes

- The tool currently focuses exclusively on **consultation-related tasks** and excludes others.  
- Assumes internal network access to the `inside.sockettelecom.com` domain and associated authentication.  
- Billing task spawning is only done for tasks classified as billable by fuzzy matching.  
- Some job type classifications are heuristic and may require tuning.  
- The `--update` CLI option is a stub for future update functionality.

---

## Development & Contribution

- The project uses synchronous Playwright API wrapped in Python classes for control and state management.  
- Logging is centralized via a custom `log_message` function writing both to file and optionally to stdout.  
- Contributions to job type classification, summary formatting, and UI interaction handling are welcome.  
- To develop locally, install dependencies and run `backend.py`. Use `--version` to check the current version.

---

## Legal & Disclaimer

This is an internal utility script for automating work order processing in the dispatch system. No warranties are provided. Use at your own risk. Internal system URLs and credentials are required for operation and are not exposed externally.

---

## Contact

For questions or support, please submit and Issue on GitHub.
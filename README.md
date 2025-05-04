# job-tracker-parser-simulator
A Dockerized Node.js environment to locally test and debug the email parsing functions (job_application_email_parser.gs) by simulating the Google Apps Script environment (GmailApp, SpreadsheetApp). Companion tool for the Automated Job Application Tracker repo.

# Job Tracker Parser Simulator (Dockerized)

## Overview

This project provides a local, Dockerized Node.js environment designed specifically for **testing and debugging the core email parsing logic** used in the main [Automated Job Application Tracker & Pipeline Manager](https://github.com/Frankwerd/Automated-Job-Application-Tracker-Pipeline-Manager/) project.

The primary goal is to simulate the Google Apps Script environment (`SpreadsheetApp`, `GmailApp`) locally, allowing you to run the parsing function (`processJobApplicationEmails`) against various predefined email scenarios without needing live Gmail/Sheet interactions or waiting for Apps Script triggers. This enables rapid testing, refinement of parsing rules (keywords, regex), and validation of how the script handles different email types (new applications, updates, rejections, interviews, offers).

## How it Works

This Node.js script:

1.  **Mocks Google Services:** It includes mock implementations of `SpreadsheetApp`, `GmailApp`, `MockSheet`, `MockThread`, and `MockMessage` classes that mimic the essential methods used by the original parser script (e.g., `getSheetByName`, `getThreads`, `getMessages`, `appendRow`, `getRange`, `addLabel`, `removeLabel`).
2.  **Contains Parser Logic:** It incorporates the core logic from the `job_application_email_parser.gs` script, adapted slightly to work within the Node.js/mock environment. This includes functions for extracting company/title, parsing status keywords, and interacting with the mock sheet/labels.
3.  **Sets Up Test Data:** It initializes a `mockSheet` array (simulating the Google Sheet) and defines functions (`addPreExistingApplications`, `createTestCases`) to populate the sheet with existing data and create mock email "threads" with various subjects, senders, bodies, and dates.
4.  **Executes the Parser:** The `runSimulation()` function orchestrates the setup and calls the adapted `processJobApplicationEmails()` function.
5.  **Outputs Results:** It logs detailed information to the console during execution, showing how emails are processed, how the mock sheet is updated or appended, and the final state of the mock sheet and labels.
6.  **Runs in Docker:** A `Dockerfile` is provided to containerize the Node.js environment and the script, ensuring consistent execution regardless of your local setup.

## Relation to the Main Project

This simulator was built as a companion tool for the [Automated Job Application Tracker & Pipeline Manager](https://github.com/Frankwerd/Automated-Job-Application-Tracker-Pipeline-Manager/).

*   It uses the **same core parsing algorithms and configuration constants** (column numbers, status keywords, label names etc.) defined in the `job_application_email_parser.gs` script from that repository.
*   It allows you to **test changes or additions to the parsing logic** (e.g., adding new keywords, improving regex) locally *before* deploying them to the live Google Apps Script environment.
*   By observing the console output, you can **verify exactly how the script interprets different emails** and updates the simulated spreadsheet, helping to debug issues found in the live tracker.

## Prerequisites

*   **Docker:** You need Docker Desktop (Windows/Mac) or Docker Engine (Linux) installed and running on your machine. Download from [docker.com](https://www.docker.com/).

## Setup & Running with Docker

Follow these steps to run the simulation locally:

1.  **Clone This Repository:**
    ```bash
    git clone [URL of THIS simulator repository]
    cd [simulator-repository-directory-name]
    ```
    *(Replace placeholders with your actual URL and directory name)*

2.  **Build the Docker Image:**
    Navigate to the repository's root directory (where the `Dockerfile` is located) in your terminal and run:
    ```bash
    docker build -t job-parser-simulator .
    ```
    *   `docker build`: The command to build an image from a Dockerfile.
    *   `-t job-parser-simulator`: Tags the image with a memorable name (`job-parser-simulator`). You can change this tag if you like.
    *   `.`: Specifies that the build context (including the `Dockerfile` and necessary code files) is the current directory.
    *   **What this does:** Docker reads the `Dockerfile`, sets up a Node.js environment inside the image, copies the necessary script files into it, and potentially installs any Node.js dependencies (if you add a `package.json` later).

3.  **Run the Simulation Container:**
    Once the image is built successfully, run the simulation using:
    ```bash
    docker run --rm -it job-parser-simulator
    ```
    *   `docker run`: The command to create and start a container from an image.
    *   `--rm`: Automatically removes the container when it exits (keeps things tidy).
    *   `-it`: Runs the container in interactive mode and allocates a pseudo-TTY, allowing you to see the `console.log` output directly in your terminal.
    *   `job-parser-simulator`: The name/tag of the image you built in the previous step.
    *   **What this does:** Starts a container based on your image. The `Dockerfile`'s `CMD` or `ENTRYPOINT` instruction (likely `node runSimulation.js` or similar) will execute the main simulation script inside the container.

## Expected Output

When you run the container, you should see output in your terminal similar to this (details will match the specific test cases):
============================================
STARTING JOB APPLICATION EMAIL PARSER SIMULATION
Mock sheet initialized with headers
SIM: Adding pre-existing application entries to mock sheet...
SIM: Added pre-existing entry for TechCorp (Senior Developer)
... (other pre-existing entries) ...
SIM: Mock sheet now has 5 rows (including header).
SIM: Creating test email threads...
SIM: Created 5 test email threads in label "JobAppToProcess"
=== RUNNING processJobApplicationEmails ===
PROCESS: Pre-loaded data for 4 companies. Found 4 existing email IDs.
PROCESS: Found 5 threads with label "JobAppToProcess".
PROCESS: ----- Processing Email ID: newapp_acme_1 -----
PROCESS: Subject: "Thank you for applying to Software Engineer at Acme Inc" From: recruiting@acme.com
PROCESS: Parsed from Subject - Company: "Acme Inc", Title: "Software Engineer"
PROCESS: NEW application entry identified for Company: "Acme Inc"
SHEET: Appended row (new length 6): Processed Timestamp, ..., Acme Inc, Software Engineer, Applied, ..., Email Subject, ..., newapp_acme_1...
PROCESS: ---> APPENDED New Row: P:Other, C:Acme Inc, T:Software Engineer, S:Applied, ID:newapp_acme_1
PROCESS: Applying label changes for thread newapp_acme_1 (Email ID: newapp_acme_1)
LABEL: Removed thread newapp_acme_1 from label "JobAppToProcess"
THREAD newapp_acme_1: Removed label "JobAppToProcess"
LABEL: Added thread newapp_acme_1 to label "TestAppToProcess"
THREAD newapp_acme_1: Added label "TestAppToProcess"
PROCESS: ----- Finished Email ID: newapp_acme_1 -----
PROCESS: ----- Processing Email ID: update_techcorp_reject_1 -----
PROCESS: Subject: "Update on your TechCorp application" From: hr@techcorp.com
PROCESS: Parsed from Subject - Company: "TechCorp", Title: "N/A - Manual Review Needed"
PROCESS: UPDATE identified for Company "TechCorp" (Found 1 matches). Targeting last entry at row: 2
PROCESS: UPDATE determined new status from body: Rejected
PROCESS: ---> UPDATED mockSheet Row 2: Status='Rejected', UpdateDate set, Email Info Updated.
PROCESS: Applying label changes for thread update_techcorp_reject_1 (Email ID: update_techcorp_reject_1)
... (similar processing logs for other emails) ...
PROCESS: Finished processing run. Added 1 new entries, updated 4 existing entries.
=== FINISHED processJobApplicationEmails ===
=== SIMULATION RESULTS ===
Final spreadsheet rows: 6
Spreadsheet Data (Company | Title | Status | Email ID | Last Update Date):
HEADERS: Processed Timestamp | Email Date | Platform | Company | Job Title | Status | Last Update Date | Email Subject | Email Link | Email ID
ROW 2: TechCorp | Senior Developer | Rejected | update_techcorp_reject_1 | 4/29/2025
ROW 3: Beta Corp | Data Analyst | Interview Scheduled | update_betacorp_interview_1 | 4/29/2025
ROW 4: Gamma Inc | Backend Engineer | Assessment/Screening | update_gammainc_assess_1 | 4/29/2025
ROW 5: Delta Solutions | Project Manager | Offer/Accepted | update_deltasol_offer_1 | 4/29/2025
ROW 6: Acme Inc | Software Engineer | Applied | newapp_acme_1 |
Final Label State:
"JobAppToProcess" Threads: 0
"TestAppToProcess" Threads: 5
============================================
SIMULATION COMPLETE
===========================================

*(Note: Exact output dates and formatting might vary slightly)*

## Customizing Test Cases

To test different email scenarios or parsing rules:

1.  **Modify Existing Cases:** Edit the email subjects, bodies, senders, or dates within the `createTestCases()` function in the main script file (e.g., `runSimulation.js` or equivalent).
2.  **Add New Cases:** Add more calls to `new MockThread(...)` within `createTestCases()`, ensuring you give them unique IDs and add them to the `mockLabels["JobAppToProcess"].threads` array.
3.  **Modify Pre-existing Data:** Adjust the entries added in the `addPreExistingApplications()` function to test updates against different initial sheet states.
4.  **Rebuild & Rerun:** After making changes to the script, you'll need to rebuild the Docker image (`docker build -t job-parser-simulator .`) and then run the container again (`docker run --rm -it job-parser-simulator`).

## Dockerfile Example

Ensure you have a `Dockerfile` in the root of this repository. A basic example:

```dockerfile
# Use an official Node.js runtime as a parent image
# Using node:18-alpine as an example; choose a version suitable for your code
FROM node:18-alpine

# Set the working directory in the container
WORKDIR /usr/src/app

# Copy the script file into the container at /usr/src/app
# !! IMPORTANT: Replace 'your_main_simulation_script.js' with the actual name of your main Node.js script file !!
COPY your_main_simulation_script.js ./
# If you have helper files in subdirectories (e.g., ./lib), copy them too:
# COPY lib ./lib

# Specify the command to run on container start
# !! IMPORTANT: Replace 'your_main_simulation_script.js' with the actual name of your main Node.js script file !!
CMD [ "node", "your_main_simulation_script.js" ]

(Remember to replace your_main_simulation_script.js with the correct filename in the COPY and CMD lines above)
License

MIT License

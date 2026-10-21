

https://github.com/user-attachments/assets/4f61fd02-95bd-4f46-bb38-0ff9f42b00dc

AI Office Agent — Powered by NVIDIA Nemotron

An AI-powered desktop automation agent that automates Microsoft Excel, Word, and PowerPoint using natural-language instructions. Instead of merely suggesting what to do, the agent executes tasks directly in the desktop applications, allowing users to observe the automation in real time.

Built with Python and powered by **NVIDIA Nemotron**, the agent combines LLM-based reasoning with tools for interacting with Microsoft Office applications.

## Overview

Repetitive office work often involves editing spreadsheets, processing documents, formatting presentations, and performing calculations. These tasks can consume valuable time when performed manually.

The AI Office Agent aims to simplify this workflow. Give it a task in natural language, and it uses its language model and specialised tools to determine the required actions and execute them in the relevant Office application.

**The key differentiator:** the agent works with the live desktop applications rather than limiting its output to text responses or standalone generated files.

## Features

### Microsoft Excel Automation

Automate spreadsheet operations, calculations, and workbook modifications.

- **Data entry:** Enter structured data into spreadsheets.
- **Workbook editing:** Modify existing Excel workbooks while working with their current contents.
- **Calculations and formulas:** Execute spreadsheet calculations and formula-based tasks.
- **Data modification:** Update existing cell values and perform requested spreadsheet operations.
- **Live execution:** Observe spreadsheet operations as the agent works in the Excel interface.

### Microsoft Word Automation

Automate document creation and editing through specialised Word tools.

- Create and modify document content.
- Edit existing text and document elements.
- Perform document operations through the available automation tools.
- Execute document-related tasks in the live Word interface.

### Microsoft PowerPoint Automation

Create and modify presentations through natural-language instructions.

- **Build presentations:** Add multiple slides in a single operation, with layouts, titles, bullet points, nested bullet points, and speaker notes.
- **Theme management:** Configure fonts, accent colours, backgrounds, and title and text colours.
- **Slide editing:** Inspect slide contents and edit text, shapes, and table cells.
- **Content insertion:** Add text boxes, pictures, and formatted tables.
- **Slide organisation:** Move, duplicate, and delete slides.
- **Export:** Save presentations and export them to PDF when requested.

### AI-Powered Task Execution

- Natural-language task instructions.
- NVIDIA Nemotron-powered language understanding and reasoning.
- Tool-based execution for specialised Office operations.
- Coordination between the LLM and Python automation functions.
- Live execution

**Setup and running guide**
 **What you need**
    Windows 10 or 11. The agent controls Office through Windows COM, so it won’t run on Mac or Linux.
    Desktop Microsoft Office (Excel, Word and PowerPoint installed and activated). Office on the web doesn’t work.
    Python 3.10 or newer.
    A Nebius Token Factory API key.

  **Install the packages**    
     openai>=1.40.0
     python-dotenv>=1.0.0
     streamlit>=1.43.0
     xlwings>=0.30.0
     pywin32>=306; sys_platform == "win32"

  **Create your .env file**
    NEBIUS_API_KEY=your_real_key
    NEBIUS_BASE_URL=https://api.tokenfactory.nebius.com/v1/
    NEBIUS_MODEL=exact_model_id

  **Find the exact model ID**
    python list_models.py (Copy the Nemotron ID it prints into NEBIUS_MODEL)

  **Start the app**
    streamlit run app.py

And you are ready to run any task on this agent...    

# Argos - Auditoria Simplificada

## Project Overview
Argos is a Python-based web application built with **Streamlit** designed to streamline auditing processes. It automates the application of audit procedures, generates reports, and integrates **Google Gemini AI** for advanced analysis of audit context and documents.

## Tech Stack
*   **Frontend/App Framework:** Streamlit
*   **Language:** Python
*   **Data Processing:** Pandas, OpenPyXL, XlsxWriter
*   **Document Generation:** Python-docx, Docxtpl
*   **AI/LLM:** Google Gen AI SDK (`google-genai`)

## Setup and Execution

### Prerequisites
*   Python 3.x
*   A valid Google Gemini API Key (for AI features)

### Installation
1.  Create a virtual environment (optional but recommended):
    ```bash
    python -m venv venv
    source venv/bin/activate  # On Linux/Mac
    # venv\Scripts\activate  # On Windows
    ```
2.  Install dependencies:
    ```bash
    pip install -r requirements.txt
    ```

### Running the Application
Execute the main Streamlit app:
```bash
streamlit run app.py
```
The application will open in your default web browser (usually at `http://localhost:8501`).

## Architecture & Key Files

### Core Structure
*   **`app.py`**: The main entry point. It initializes session state variables (e.g., `audit_results`, `files_processed`) and handles the multi-page navigation logic.
*   **`classes.py`**: Defines the domain model and business logic:
    *   `Auditado`: Represents the entity being audited.
    *   `ProcedimentoAuditoria`: A set of verification actions.
    *   `AcaoVerificacao`: Atomic checks performed on data sources.
    *   `Achado`: Represents a finding when a procedure identifies an issue.
    *   `FonteInformacao`: Wraps data sources (Excel files).
*   **`utils.py`**: Contains utility functions for:
    *   Interacting with the Gemini API (`avalia_gemini`).
    *   Processing Word/Excel templates.
    *   Logic expression evaluation (`avalia_expressao`).
    *   Jinja2 template rendering.

### Directories
*   **`pages/`**: Individual Streamlit pages corresponding to the app's workflow steps:
    *   `home.py`: Landing page.
    *   `aplica_procedimentos.py`: Logic to run audit procedures.
    *   `analise_gemini.py`: Interface for AI-powered analysis of audit documents.
    *   `escreve_relatorio.py`: Generates the final audit report.
*   **`prompts/`**: Contains Markdown files with Jinja2 templates used as prompts for the Gemini AI.
*   **`docs/`**: Stores Word (`.docx`) and Excel (`.xlsx`) templates used for report generation.

## Development Conventions

### State Management
The app relies heavily on `st.session_state` to pass data between pages. Key state variables include:
*   `audit_results`: Stores the dictionary of `Auditado` objects and results.
*   `audit_completed`: Boolean flag indicating if the audit logic has run.
*   `files_processed`: Boolean flag for file upload status.

### AI Integration
*   **Client:** Uses `google.genai.Client`.
*   **Prompts:** Stored in `prompts/` and rendered using **Jinja2**, allowing injection of audit data (`{{ auditado.nome }}`, etc.) directly into the prompt context.
*   **Context:** Supports uploading files (PDF, Excel, ZIP) which are uploaded to the Gemini API via `client.files.upload`.
*   **Output:** Supports both free-text and structured JSON output (parsed into Pandas DataFrames).

### Auditing Logic (`classes.py`)
The core logic follows this hierarchy:
1.  **Auditado** applies a **ProcedimentoAuditoria**.
2.  A **ProcedimentoAuditoria** executes multiple **AcaoVerificacao** items.
3.  Each **AcaoVerificacao** checks a specific field in a **FonteInformacao** against a condition (`situacao_inconforme`).
4.  If the conditions defined in `logica_achado` are met, an **Achado** is created and attached to the **Auditado**.

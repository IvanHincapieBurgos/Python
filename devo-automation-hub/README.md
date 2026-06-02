# Devo Automation Hub

A comprehensive Python automation toolkit for managing orders and tracking shipments. This project provides robust web automation, data logging, and scheduling capabilities for enterprise workflow management.

## Overview

The Devo Automation Hub is designed to streamline repetitive business operations through:

- **Web Automation**: Selenium-based automation for browser interactions and order management
- **Google Sheets Integration**: Read/write operations and real-time data synchronization
- **Google Drive Logging**: Automatic logging of actions and errors to cloud storage
- **Task Scheduling**: Background schedulers for periodic operations
- **Chrome Driver Management**: Intelligent Chrome WebDriver initialization and management
- **NLP Capabilities**: Sentence transformers for advanced text processing

## Project Structure

```
devo-automation-hub/
├── devos.py                     # Main automation module with core functions
├── requirements.txt             # Python dependencies
├── credencialesdevo.json       # Google API credentials (not included in repo)
└── registros/                   # Logging and registration module
    ├── logger_helper.py         # Helper for logging actions and errors
    ├── __init__.py              # Package initialization
    └── README.md                # Registros module documentation
```

## Key Features

### 1. Web Automation & Browser Control

- **Chrome Driver Management**: Automatic initialization, debugging, and configuration
- **Selenium Integration**: Full Selenium WebDriver support for complex web interactions
- **Session Management**: Handle Chrome debug protocols and remote debugging
- **Custom Headers**: Inject custom headers for API requests via Chrome DevTools Protocol

### 2. Google Sheets Integration

- **Safe Cell Updates**: Retry logic for robust Google Sheets API interactions
- **Batch Operations**: Write multiple cells efficiently
- **Data Synchronization**: Real-time updates to spreadsheets

### 3. Google Drive Logging

The `registros` module provides centralized logging:

- **Action Logging**: Track successful operations
- **Error Logging**: Capture and store error details with stack traces
- **Automatic Upload**: CSV files automatically uploaded to Google Drive
- **Caller Detection**: Automatically detects calling function and line number

### 4. Task Scheduling

- **Background Schedulers**: APScheduler for running tasks at scheduled intervals
- **Cron Triggers**: Support for cron-like scheduling expressions
- **Timezone Support**: Proper timezone handling for global operations

### 5. Order & Tracking Management

- **Label Processing**: Download and manage shipping labels
- **Tracking Integration**: Retrieve and update shipment tracking information
- **Order Status Updates**: Track order progression through the system
- **Secondary Links Management**: Handle backup/alternative shipping links

## Dependencies

### Core Libraries

- **selenium** (≥4.10.0): Web automation framework
- **gspread** (≥5.12.0): Google Sheets Python API wrapper
- **pandas** (≥2.1.0): Data manipulation and analysis
- **APScheduler** (≥3.10.1): Advanced scheduling library
- **Pillow** (≥10.1.0): Image processing
- **PyPDF2** (≥3.0.0): PDF manipulation

### Google Cloud Integration

- **google-api-python-client** (≥2.102.0): Google API client library
- **google-auth** (≥2.26.0): Authentication library
- **google-auth-httplib2** (≥0.2.0): HTTP transport
- **google-auth-oauthlib** (≥1.2.0): OAuth2 flow

### Web & API Tools

- **requests** (≥2.31.0): HTTP library
- **beautifulsoup4** (≥4.12.2): HTML/XML parsing
- **curl_cffi** (≥0.9.0): Advanced HTTP requests with cURL
- **webdriver-manager** (≥4.0.0): Automatic WebDriver management

### Additional Tools

- **python-dotenv** (≥1.0.0): Environment variable management
- **slack_bolt** (≥1.18.0): Slack bot framework
- **psycopg2-binary** (≥2.9.9): PostgreSQL adapter
- **openpyxl** (≥3.1.0): Excel file manipulation
- **sentence-transformers**: NLP embeddings and similarity

See `requirements.txt` for complete dependency list with pinned versions.

## Installation

### Prerequisites

- Python 3.8+
- pip or conda
- Google API credentials (JSON file)
- Chrome/Chromium browser (for Selenium)

### Setup Steps

1. **Clone the repository** (or download the project)

2. **Create a virtual environment** (recommended):

3. **Install dependencies**:
   ```bash
   pip install -r requirements.txt
   ```
   
   The project includes automatic dependency installation on first run if any packages are missing.

4. **Configure Google API credentials**:
   - Create `credencialesdevo.json` with your Google service account credentials
   - Place it in the root project directory
   - For the registros module, create `credenciales_registros.json` in the `registros/` folder

5. **Set up environment variables** (optional):
   Create a `.env` file for sensitive configuration:
   ```
   GOOGLE_DRIVE_FOLDER_ID=
   TIMEZONE=America/New_York
   ```

## Configuration

### Chrome Driver Configuration

The project automatically manages Chrome WebDriver through several utility functions:

- `delete_wdm_and_prepare_service()`: Clean WebDriver Manager cache and prepare service
- `_build_return_chrome_options()`: Build optimized Chrome options
- `is_chrome_running()`: Check if Chrome is running on debug port
- `launch_chrome_window()`: Launch Chrome in debug mode

### Logging Configuration

Create a `credenciales_registros.json` with Google Drive API credentials:

## Registros Module

The `registros` subfolder contains logging utilities:

### Main Functions

- **`registrar_accion(accion, detalles=None)`**: Log successful operations to Google Drive
- **`registrar_error(mensaje, excepcion=None)`**: Log errors with full traceback


See [registros/README.md](registros/README.md) for detailed documentation.

## Architecture & Design Patterns

### Automatic Dependency Installation

The project includes a setup routine that:
1. Reads `requirements.txt`
2. Checks for missing packages
3. Automatically installs missing dependencies
4. Maps package names to importable module names

### Web Automation Flow

```
Initialize Chrome Driver
    ↓
Ensure Debug Window
    ↓
Login to OMS
    ↓
Execute Automation Tasks
    ↓
Log Actions/Errors
    ↓
Close Driver
```

### Google Sheets Integration Flow

```
Authenticate with Google Sheets API
    ↓
Open Spreadsheet
    ↓
Read/Write Data
    ↓
Retry on Failure
    ↓
Update Google Drive
```

## Troubleshooting

### Missing Dependencies

If you see `ModuleNotFoundError`:
1. Ensure `requirements.txt` is in the project root
2. Run the script again - it will auto-install dependencies
3. Or manually install: `pip install -r requirements.txt`

### Chrome Driver Issues

- **Port already in use**: Change debug port or kill existing Chrome process
- **WebDriver not found**: Clear cache: `python -c "from devos import delete_wdm_and_prepare_service; delete_wdm_and_prepare_service()"`
- **Timeout errors**: Increase timeout values in driver initialization

### Google Sheets Authorization

- Verify `credencialesdevo.json` exists and is valid
- Ensure service account has editor access to target spreadsheets
- Check that API is enabled in Google Cloud Console

### Logging Issues

- Confirm `credenciales_registros.json` exists in `registros/` folder
- Verify Google Drive API is enabled
- Check that logging folders exist in Google Drive

## Performance Tips

1. **Batch Operations**: Use `sheets_row_write()` for multiple cells instead of repeated `safe_update_cell()` calls
2. **Connection Pooling**: The `curl_cffi` library provides efficient HTTP connection handling
3. **Scheduler Optimization**: Use appropriate cron intervals to avoid overloading
4. **Error Handling**: Implement retry logic for transient failures

## Security Considerations

- Store credentials in environment variables, not in code
- Use `.gitignore` to exclude `*.json` credential files
- Implement proper authentication flows for Google APIs
- Validate and sanitize user inputs in web automation
- Use HTTPS for all external API calls
- Keep dependencies updated for security patches
# HR Automation Tool

The **HR Automation Tool** is a Django-based application designed to automate PowerPoint presentations for employee work anniversaries and manage timesheet tasks. This project includes a web interface, backend logic, and scripts for handling various HR-related automation tasks.

---

## Features

1. **Anniversary Automation**:
   - Automatically generate PowerPoint presentations for employee work anniversaries.
   - Customizable templates for branding and personalization.

2. **Timesheet Validation**:
   - Validate employee timesheets for errors.
   - Archive and manage validated timesheets.

3. **Web Interface**:
   - User-friendly interface for accessing tools.
   - Sidebar navigation for quick access to features.

---

## Installation

1. Clone the repository:
   ```bash
   git clone <repository-url>
   cd HR-automation-tool

2. Create a virtual environment:
   ```bash
   python -m venv env
   source env/bin/activate

3. Install dependencies:
   ```bash
   pip install -r requirements.txt

4. Apply database migrations:    
   ```bash
   python manage.py migrate

5. Run the development server:
   ```bash
   python manage.py runserver

---   

## Usage

1. Access the web interface at http://127.0.0.1:8000/.
2. Use the sidebar to navigate between tools:
   - PPT Automation: Generate anniversary presentations.
   - Timesheet Validation: Validate and manage timesheets.

---

## Configuration

- Static Files: Static files are served from the staticfiles/ directory. Use python manage.py collectstatic to collect static files.
- Media Files: Uploaded and generated files are stored in the media/ directory. Update MEDIA_ROOT and MEDIA_URL in settings.py if needed.

---

## Scripts
**Timesheet Validation**
 The scripts/timesheet_validation.py script handles timesheet validation tasks. It:
   - Validates timesheets for errors.
   - Archives validated timesheets.
   - Generates summary reports.


   

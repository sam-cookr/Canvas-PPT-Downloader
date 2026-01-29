# Canvas PowerPoint Downloader

A tool to download all PowerPoint files from your Canvas course modules for revision. Available as both a user-friendly web interface and a command-line script.

## Features

- Download PowerPoint files from Canvas courses (.ppt, .pptx, .pptm)
- Web interface - no command line needed
- Select specific modules or download all at once
- Real-time progress updates
- Files organized by module name
- Download as a single ZIP file (web app)
- Works with both direct files and files embedded in pages

## Installation

Install the required dependencies:
```bash
pip install -r requirements.txt
```

## Usage

### Option 1: Web App (Recommended)

The easiest way to use this tool is through the web interface:

1. **Start the web app**:
   ```bash
   streamlit run app.py
   ```

2. **Your browser will automatically open** to `http://localhost:8501`

3. **Follow the steps in the web interface**:
   - Enter your Canvas URL (e.g., `newcastle.instructure.com`)
   - Enter your API token
   - Click "Fetch My Courses"
   - Select your course from the dropdown
   - Click "Load Modules"
   - Choose which modules to download (or select all)
   - Click "Download PowerPoints"
   - Download the ZIP file with all your files

### Option 2: Command Line

If you prefer the command line:

```bash
python download_canvas_ppts.py
```

You'll be prompted for:
- Canvas URL
- API token
- Course ID

## Getting Your Credentials

### Canvas URL
Your institution's Canvas URL (e.g., `newcastle.instructure.com`)

### API Token
1. Log in to Canvas
2. Go to Account → Settings
3. Scroll to "Approved Integrations"
4. Click "+ New Access Token"
5. Copy the generated token

### Course ID (Command Line Only)

The web app shows all your courses automatically. For the command line version:

**Method 1: From the URL**
- Go to your course in Canvas
- Look at the URL: `https://newcastle.instructure.com/courses/12345`
- The course ID is `12345`

**Method 2: List all courses**
Run this command to see all your course IDs:
```python
import requests

CANVAS_URL = "https://newcastle.instructure.com"
API_TOKEN = "your_token_here"

headers = {'Authorization': f'Bearer {API_TOKEN}'}
response = requests.get(f"{CANVAS_URL}/api/v1/courses", headers=headers, params={'per_page': 100})

for course in response.json():
    if 'name' in course:
        print(f"{course['id']}: {course['name']}")
```

## How It Works

1. Connects to Canvas using your API token
2. Fetches all modules in your course
3. Scans each module for PowerPoint files (.ppt, .pptx, .pptm)
4. Downloads them organized by module name
5. Saves to `canvas_powerpoints` folder (CLI) or ZIP file (web app)

## Output Structure

```
canvas_powerpoints/
├── Week 1 - Introduction/
│   ├── Lecture_1.pptx
│   └── Tutorial_1.pptx
├── Week 2 - Data Structures/
│   ├── Lecture_2.pptx
│   └── Lab_2.pptx
└── ...
```

## Troubleshooting

**"ModuleNotFoundError: No module named 'streamlit'"**
- Run: `pip install -r requirements.txt`

**"Address already in use" (web app)**
- Another instance is running. Close it or use a different port:
  ```bash
  streamlit run app.py --server.port 8502
  ```

**"Authentication failed"**
- Double-check your API token is correct and hasn't expired
- Generate a new token if needed

**"Course not found"**
- Verify the course ID is correct (CLI)
- Make sure you're enrolled in the course

**"No courses found" (web app)**
- Check your Canvas URL is correct (don't include `https://`)
- Verify your API token is valid

**No files downloaded**
- Check that the course has PowerPoint files in modules
- Some files might be external links rather than uploaded files

## Notes

- Only downloads PowerPoint files (not PDFs or other formats)
- Files are organized by module name
- Running again will overwrite existing files
- API token is kept secure and only used for authentication
- Web app hides your token (shows as `****`) for security
- To stop the web app, press `Ctrl+C` in the terminal

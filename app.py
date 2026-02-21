#!/usr/bin/env python3
"""
Canvas File Downloader - Streamlit Web App
Downloads PowerPoint and PDF files from Canvas course modules
"""

import streamlit as st
import requests
import re
from pathlib import Path
from html.parser import HTMLParser
import zipfile
import io
import tempfile
import shutil

# ── Known Canvas institutions ──────────────────────────────────────────────────
# URLs verified against known Canvas deployments. Select "Other" if yours isn't listed.
UNIVERSITIES = {
    "Newcastle University":              "https://ncl.instructure.com",
    "Brunel University London":          "https://brunel.instructure.com",
    "Coventry University":               "https://coventry.instructure.com",
    "De Montfort University":            "https://dmu.instructure.com",
    "Leeds Beckett University":          "https://leedsbeckett.instructure.com",
    "University of Hertfordshire":       "https://herts.instructure.com",
    "University of Lincoln":             "https://lincoln.instructure.com",
    "University of Roehampton":          "https://roehampton.instructure.com",
    "University of the Arts London":     "https://canvas.arts.ac.uk",
    "RMIT University":                   "https://rmit.instructure.com",
    "University of Sydney":              "https://canvas.sydney.edu.au",
    "Cornell University":                "https://canvas.cornell.edu",
    "Indiana University":                "https://iu.instructure.com",
    "UC Berkeley":                       "https://bcourses.berkeley.edu",
    "University of Michigan":            "https://umich.instructure.com",
    "University of Washington":          "https://canvas.uw.edu",
    "Other (enter manually)":            None,
}


class FileURLExtractor(HTMLParser):
    """Extract file URLs from HTML content"""

    def __init__(self):
        super().__init__()
        self.file_urls = []

    def handle_starttag(self, tag, attrs):
        if tag == 'a':
            for attr, value in attrs:
                if attr == 'href' and value:
                    if '/files/' in value or value.endswith(('.ppt', '.pptx', '.pptm', '.pdf')):
                        self.file_urls.append(value)


def get_courses(canvas_url, api_token):
    """Fetch all courses for the user"""
    headers = {'Authorization': f'Bearer {api_token}'}
    url = f"{canvas_url}/api/v1/courses"
    params = {'per_page': 100, 'enrollment_state': 'active'}

    try:
        response = requests.get(url, headers=headers, params=params)
        response.raise_for_status()
        courses = response.json()
        return [c for c in courses if 'name' in c and 'id' in c]
    except Exception as e:
        st.error(f"Error fetching courses: {e}")
        return []


def get_modules(canvas_url, api_token, course_id):
    """Fetch all modules in a course"""
    headers = {'Authorization': f'Bearer {api_token}'}
    url = f"{canvas_url}/api/v1/courses/{course_id}/modules"
    params = {'per_page': 100}

    modules = []
    while url:
        response = requests.get(url, headers=headers, params=params)
        response.raise_for_status()
        modules.extend(response.json())

        if 'next' in response.links:
            url = response.links['next']['url']
            params = {}
        else:
            url = None

    return modules


def get_module_items(canvas_url, api_token, course_id, module_id):
    """Fetch all items in a module"""
    headers = {'Authorization': f'Bearer {api_token}'}
    url = f"{canvas_url}/api/v1/courses/{course_id}/modules/{module_id}/items"
    params = {'per_page': 100}

    items = []
    while url:
        response = requests.get(url, headers=headers, params=params)
        response.raise_for_status()
        items.extend(response.json())

        if 'next' in response.links:
            url = response.links['next']['url']
            params = {}
        else:
            url = None

    return items


def get_page_content(api_token, page_url):
    """Fetch the content of a Canvas page"""
    headers = {'Authorization': f'Bearer {api_token}'}
    try:
        response = requests.get(page_url, headers=headers)
        response.raise_for_status()
        page_data = response.json()
        return page_data.get('body', '')
    except:
        return ''


def extract_file_id_from_url(url):
    """Extract Canvas file ID from URL"""
    match = re.search(r'/files/(\d+)', url)
    if match:
        return match.group(1)
    return None


def get_file_info(canvas_url, api_token, course_id, file_id):
    """Get file information from Canvas API"""
    headers = {'Authorization': f'Bearer {api_token}'}
    try:
        url = f"{canvas_url}/api/v1/courses/{course_id}/files/{file_id}"
        response = requests.get(url, headers=headers)
        response.raise_for_status()
        return response.json()
    except:
        return None


def download_file(api_token, url):
    """Download a file and return its content"""
    headers = {'Authorization': f'Bearer {api_token}'}
    response = requests.get(url, headers=headers, stream=True)
    response.raise_for_status()
    return response.content


def is_downloadable_file(filename):
    """Check if file is a PowerPoint or PDF"""
    return filename.lower().endswith(('.ppt', '.pptx', '.pptm', '.pdf'))


def sanitize_filename(filename):
    """Remove invalid characters from filename"""
    invalid_chars = '<>:"/\\|?*'
    for char in invalid_chars:
        filename = filename.replace(char, '_')
    return filename


def download_powerpoints(canvas_url, api_token, course_id, selected_module_ids, progress_bar, status_text):
    """Download PowerPoints and PDFs from selected modules"""
    headers = {'Authorization': f'Bearer {api_token}'}

    # Create temporary directory
    temp_dir = Path(tempfile.mkdtemp())

    try:
        modules = get_modules(canvas_url, api_token, course_id)
        selected_modules = [m for m in modules if m['id'] in selected_module_ids]

        total_files = 0
        downloaded_files = set()

        for idx, module in enumerate(selected_modules):
            module_name = module['name']
            module_id = module['id']

            status_text.markdown(f'<p class="status-text">Processing module: <strong>{module_name}</strong></p>', unsafe_allow_html=True)

            # Create subfolder
            module_folder = temp_dir / sanitize_filename(module_name)
            module_folder.mkdir(exist_ok=True)

            # Get module items
            items = get_module_items(canvas_url, api_token, course_id, module_id)

            for item in items:
                item_type = item.get('type')
                item_title = item.get('title', 'unknown')

                # Handle direct File items
                if item_type == 'File':
                    filename = item_title

                    if is_downloadable_file(filename):
                        file_url = item.get('url')
                        if file_url:
                            file_response = requests.get(file_url, headers=headers)
                            file_data = file_response.json()
                            download_url = file_data.get('url')

                            if download_url and download_url not in downloaded_files:
                                status_text.markdown(f'<p class="status-text">Downloading: <strong>{filename}</strong></p>', unsafe_allow_html=True)
                                content = download_file(api_token, download_url)
                                filepath = module_folder / sanitize_filename(filename)
                                filepath.write_bytes(content)
                                downloaded_files.add(download_url)
                                total_files += 1

                # Handle Page items
                elif item_type == 'Page':
                    page_url = item.get('url')

                    if page_url:
                        page_html = get_page_content(api_token, page_url)
                        parser = FileURLExtractor()
                        parser.feed(page_html)

                        for file_url in parser.file_urls:
                            file_id = extract_file_id_from_url(file_url)

                            if file_id:
                                file_info = get_file_info(canvas_url, api_token, course_id, file_id)

                                if file_info:
                                    filename = file_info.get('display_name', file_info.get('filename', 'unknown'))

                                    if is_downloadable_file(filename):
                                        download_url = file_info.get('url')

                                        if download_url and download_url not in downloaded_files:
                                            status_text.markdown(f'<p class="status-text">Downloading: <strong>{filename}</strong></p>', unsafe_allow_html=True)
                                            content = download_file(api_token, download_url)
                                            filepath = module_folder / sanitize_filename(filename)
                                            filepath.write_bytes(content)
                                            downloaded_files.add(download_url)
                                            total_files += 1

            # Update progress
            progress_bar.progress((idx + 1) / len(selected_modules))

        # Create ZIP file
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
            for file_path in temp_dir.rglob('*'):
                if file_path.is_file():
                    arcname = file_path.relative_to(temp_dir)
                    zip_file.write(file_path, arcname)

        zip_buffer.seek(0)
        return zip_buffer, total_files

    finally:
        # Clean up temp directory
        shutil.rmtree(temp_dir, ignore_errors=True)


def inject_css():
    st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Cormorant+Garamond:ital,wght@0,400;0,600;0,700;1,400;1,600&family=DM+Sans:opsz,wght@9..40,300;9..40,400;9..40,500&family=JetBrains+Mono:wght@400;500&display=swap');

    :root {
        --bg-deep:       #0d0d15;
        --bg-surface:    #13131f;
        --bg-elevated:   #1b1b2c;
        --bg-card:       #1e1e30;
        --gold:          #c9a84c;
        --gold-bright:   #deba6a;
        --gold-dim:      rgba(201, 168, 76, 0.12);
        --gold-glow:     rgba(201, 168, 76, 0.22);
        --blue:          #5c8ff0;
        --blue-dim:      rgba(92, 143, 240, 0.12);
        --text-primary:  #ede8dc;
        --text-secondary:#7c7c92;
        --text-muted:    #3e3e54;
        --border:        rgba(201, 168, 76, 0.10);
        --border-strong: rgba(201, 168, 76, 0.28);
        --success-color: #5dba7d;
        --error-color:   #e07070;
        --warn-color:    #e0a050;
    }

    /* ── Global reset ───────────────────────────────── */
    .stApp {
        background-color: var(--bg-deep) !important;
        background-image:
            radial-gradient(ellipse 60% 40% at 15% -5%,  rgba(92,143,240,0.07) 0%, transparent 60%),
            radial-gradient(ellipse 50% 35% at 88% 105%, rgba(201,168,76,0.07) 0%, transparent 60%);
        font-family: 'DM Sans', sans-serif !important;
        color: var(--text-primary) !important;
    }

    #MainMenu, footer, [data-testid="stToolbar"] { visibility: hidden; }

    [data-testid="stHeader"] {
        background: transparent !important;
        backdrop-filter: none !important;
        height: 0 !important;
        min-height: 0 !important;
    }

    .block-container {
        padding-top: 2.5rem !important;
        padding-bottom: 4rem !important;
        max-width: 760px !important;
    }

    /* ── Typography defaults ────────────────────────── */
    p, li, span, div {
        font-family: 'DM Sans', sans-serif;
        color: var(--text-secondary);
    }

    /* ── Hero ───────────────────────────────────────── */
    .hero {
        padding: 3rem 0 2.2rem;
        border-bottom: 1px solid var(--border);
        margin-bottom: 0.25rem;
    }
    .hero-eyebrow {
        font-family: 'JetBrains Mono', monospace;
        font-size: 0.62rem;
        letter-spacing: 0.22em;
        text-transform: uppercase;
        color: var(--gold);
        opacity: 0.65;
        margin-bottom: 0.7rem;
    }
    .hero-title {
        font-family: 'Cormorant Garamond', serif;
        font-size: 3.6rem;
        font-weight: 600;
        color: var(--text-primary);
        line-height: 1.0;
        letter-spacing: -0.03em;
        margin-bottom: 0.9rem;
    }
    .hero-title em {
        font-style: italic;
        color: var(--gold-bright);
    }
    .hero-sub {
        font-family: 'DM Sans', sans-serif;
        font-size: 0.88rem;
        color: var(--text-secondary);
        line-height: 1.65;
        max-width: 500px;
    }

    /* ── Step headers ───────────────────────────────── */
    .step-block {
        display: flex;
        align-items: center;
        gap: 14px;
        margin: 2.8rem 0 1.2rem;
        padding-bottom: 0.75rem;
        border-bottom: 1px solid var(--border);
    }
    .step-num {
        width: 26px;
        height: 26px;
        flex-shrink: 0;
        border-radius: 50%;
        border: 1px solid var(--border-strong);
        background: var(--gold-dim);
        display: flex;
        align-items: center;
        justify-content: center;
        font-family: 'JetBrains Mono', monospace;
        font-size: 0.62rem;
        color: var(--gold);
        letter-spacing: 0.02em;
    }
    .step-label {
        font-family: 'Cormorant Garamond', serif;
        font-size: 1.35rem;
        font-weight: 600;
        color: var(--text-primary);
        letter-spacing: -0.01em;
        line-height: 1;
    }
    .step-label small {
        font-family: 'DM Sans', sans-serif;
        font-size: 0.72rem;
        font-weight: 400;
        color: var(--text-muted);
        letter-spacing: 0.04em;
        margin-left: 10px;
        vertical-align: middle;
    }

    /* ── Form labels ────────────────────────────────── */
    [data-testid="stTextInput"] label p,
    [data-testid="stSelectbox"] label p,
    [data-testid="stMultiSelect"] label p,
    [data-testid="stCheckbox"] label p {
        font-family: 'JetBrains Mono', monospace !important;
        font-size: 0.65rem !important;
        letter-spacing: 0.15em !important;
        text-transform: uppercase !important;
        color: var(--gold) !important;
        opacity: 0.75 !important;
    }

    /* ── Text inputs ────────────────────────────────── */
    [data-testid="stTextInput"] input {
        background: var(--bg-elevated) !important;
        border: 1px solid var(--border) !important;
        border-radius: 4px !important;
        color: var(--text-primary) !important;
        font-family: 'JetBrains Mono', monospace !important;
        font-size: 0.83rem !important;
        padding: 10px 14px !important;
        transition: border-color 0.2s, box-shadow 0.2s !important;
        caret-color: var(--gold) !important;
    }
    [data-testid="stTextInput"] input:focus {
        border-color: var(--border-strong) !important;
        box-shadow: 0 0 0 3px var(--gold-dim) !important;
        outline: none !important;
    }
    [data-testid="stTextInput"] input::placeholder {
        color: var(--text-muted) !important;
        opacity: 1 !important;
    }

    /* ── Selectbox ──────────────────────────────────── */
    [data-testid="stSelectbox"] > div > div,
    [data-testid="stSelectbox"] [data-baseweb="select"] > div {
        background: var(--bg-elevated) !important;
        border: 1px solid var(--border) !important;
        border-radius: 4px !important;
        color: var(--text-primary) !important;
        font-family: 'DM Sans', sans-serif !important;
    }
    [data-testid="stSelectbox"] [data-baseweb="select"] > div:focus-within {
        border-color: var(--border-strong) !important;
        box-shadow: 0 0 0 3px var(--gold-dim) !important;
    }
    [data-testid="stSelectbox"] span {
        color: var(--text-primary) !important;
    }

    /* ── Multiselect ────────────────────────────────── */
    [data-testid="stMultiSelect"] [data-baseweb="select"] > div {
        background: var(--bg-elevated) !important;
        border: 1px solid var(--border) !important;
        border-radius: 4px !important;
    }
    [data-testid="stMultiSelect"] [data-baseweb="tag"] {
        background: var(--gold-dim) !important;
        border: 1px solid var(--border-strong) !important;
        border-radius: 3px !important;
    }
    [data-testid="stMultiSelect"] [data-baseweb="tag"] span {
        color: var(--gold-bright) !important;
        font-family: 'JetBrains Mono', monospace !important;
        font-size: 0.7rem !important;
    }

    /* ── Buttons ────────────────────────────────────── */
    .stButton > button {
        background: transparent !important;
        border: 1px solid var(--border-strong) !important;
        color: var(--gold) !important;
        font-family: 'JetBrains Mono', monospace !important;
        font-size: 0.72rem !important;
        letter-spacing: 0.1em !important;
        text-transform: uppercase !important;
        border-radius: 3px !important;
        padding: 9px 22px !important;
        transition: all 0.18s ease !important;
        cursor: pointer !important;
    }
    .stButton > button:hover:not(:disabled) {
        background: var(--gold-dim) !important;
        border-color: var(--gold) !important;
        box-shadow: 0 2px 12px var(--gold-dim) !important;
    }
    .stButton > button:disabled {
        opacity: 0.35 !important;
        cursor: not-allowed !important;
    }
    /* Primary buttons */
    .stButton > button[kind="primary"],
    [data-testid="baseButton-primary"] {
        background: var(--gold) !important;
        border-color: var(--gold) !important;
        color: var(--bg-deep) !important;
        font-weight: 600 !important;
    }
    .stButton > button[kind="primary"]:hover:not(:disabled),
    [data-testid="baseButton-primary"]:hover:not(:disabled) {
        background: var(--gold-bright) !important;
        border-color: var(--gold-bright) !important;
        transform: translateY(-1px) !important;
        box-shadow: 0 4px 18px var(--gold-glow) !important;
    }

    /* ── Download button ────────────────────────────── */
    [data-testid="stDownloadButton"] button {
        background: var(--blue) !important;
        border: 1px solid var(--blue) !important;
        color: #fff !important;
        font-family: 'JetBrains Mono', monospace !important;
        font-size: 0.72rem !important;
        letter-spacing: 0.1em !important;
        text-transform: uppercase !important;
        border-radius: 3px !important;
        padding: 10px 26px !important;
        transition: all 0.18s ease !important;
    }
    [data-testid="stDownloadButton"] button:hover {
        background: #7aaaf5 !important;
        transform: translateY(-1px) !important;
        box-shadow: 0 4px 18px var(--blue-dim) !important;
    }

    /* ── Checkbox ───────────────────────────────────── */
    [data-testid="stCheckbox"] label {
        color: var(--text-secondary) !important;
    }
    [data-testid="stCheckbox"] [data-baseweb="checkbox"] > div:first-child {
        background: var(--bg-elevated) !important;
        border-color: var(--border-strong) !important;
    }
    [data-testid="stCheckbox"] [data-baseweb="checkbox"] input:checked + div {
        background: var(--gold) !important;
        border-color: var(--gold) !important;
    }

    /* ── Progress bar ───────────────────────────────── */
    [data-testid="stProgress"] > div {
        background: var(--bg-elevated) !important;
        border-radius: 2px !important;
    }
    [data-testid="stProgress"] > div > div {
        background: linear-gradient(90deg, var(--gold) 0%, var(--gold-bright) 100%) !important;
        border-radius: 2px !important;
    }

    /* ── Alert boxes ────────────────────────────────── */
    [data-testid="stAlert"] {
        background: var(--bg-elevated) !important;
        border-radius: 4px !important;
    }
    [data-testid="stAlert"] p {
        color: var(--text-primary) !important;
        font-family: 'DM Sans', sans-serif !important;
        font-size: 0.87rem !important;
    }
    /* Success */
    [data-testid="stAlert"][data-baseweb="notification"][kind="success"],
    div[data-baseweb="notification"]:has(svg[data-icon="check-circle"]) {
        border-left: 3px solid var(--success-color) !important;
    }
    /* Error */
    [data-testid="stAlert"][data-baseweb="notification"][kind="error"] {
        border-left: 3px solid var(--error-color) !important;
    }
    /* Warning */
    [data-testid="stAlert"][data-baseweb="notification"][kind="warning"] {
        border-left: 3px solid var(--warn-color) !important;
    }
    /* Info */
    [data-testid="stAlert"][data-baseweb="notification"][kind="info"] {
        border-left: 3px solid var(--blue) !important;
    }

    /* ── Spinner ────────────────────────────────────── */
    [data-testid="stSpinner"] > div {
        border-top-color: var(--gold) !important;
    }

    /* ── Divider ────────────────────────────────────── */
    hr { border-color: var(--border) !important; margin: 2.5rem 0 !important; }

    /* ── Status text (during download) ─────────────── */
    .status-text {
        font-family: 'JetBrains Mono', monospace !important;
        font-size: 0.72rem !important;
        color: var(--text-secondary) !important;
        letter-spacing: 0.04em;
    }
    .status-text strong { color: var(--gold) !important; font-weight: 500 !important; }

    /* ── Count badge ────────────────────────────────── */
    .count-badge {
        display: inline-flex;
        align-items: center;
        gap: 7px;
        background: var(--gold-dim);
        border: 1px solid var(--border-strong);
        border-radius: 3px;
        padding: 5px 12px;
        font-family: 'JetBrains Mono', monospace;
        font-size: 0.68rem;
        color: var(--gold);
        letter-spacing: 0.08em;
        margin: 0.4rem 0;
    }

    /* ── Tip card ───────────────────────────────────── */
    .tip-card {
        background: var(--bg-elevated);
        border: 1px solid var(--border);
        border-radius: 4px;
        padding: 14px 18px;
        font-family: 'DM Sans', sans-serif;
        font-size: 0.80rem;
        color: var(--text-secondary);
        line-height: 1.65;
        margin-top: 2.5rem;
    }
    .tip-card strong { color: var(--gold); font-weight: 500; }
    .tip-card code {
        font-family: 'JetBrains Mono', monospace;
        font-size: 0.75rem;
        background: var(--gold-dim);
        padding: 1px 5px;
        border-radius: 2px;
        color: var(--gold-bright);
    }

    /* ── Columns gap ────────────────────────────────── */
    [data-testid="stHorizontalBlock"] { gap: 16px !important; }
    </style>
    """, unsafe_allow_html=True)


def step_header(number, title, subtitle=""):
    sub = f'<small>{subtitle}</small>' if subtitle else ''
    st.markdown(f"""
    <div class="step-block">
        <div class="step-num">{number:02d}</div>
        <div class="step-label">{title}{sub}</div>
    </div>
    """, unsafe_allow_html=True)


def main():
    st.set_page_config(
        page_title="Canvas Downloader",
        page_icon="⬇",
        layout="centered"
    )

    inject_css()

    # ── Hero ──────────────────────────────────────────────
    st.markdown("""
    <div class="hero">
        <div class="hero-eyebrow">Canvas LMS · File Extraction Tool</div>
        <div class="hero-title">Canvas <em>File</em><br>Downloader</div>
        <div class="hero-sub">
            Bulk-download PowerPoint and PDF files from your Canvas
            course modules — packaged into a single ZIP archive.
        </div>
    </div>
    """, unsafe_allow_html=True)

    # ── Session state ──────────────────────────────────────
    if 'courses' not in st.session_state:
        st.session_state.courses = []
    if 'modules' not in st.session_state:
        st.session_state.modules = []

    # ── Step 1 · Configuration ─────────────────────────────
    step_header(1, "Configuration", "Select your university & enter API credentials")

    uni_names = list(UNIVERSITIES.keys())
    col1, col2 = st.columns(2)
    with col1:
        selected_uni = st.selectbox(
            "University",
            options=uni_names,
            index=uni_names.index("Newcastle University"),
            help="Choose your institution. Select 'Other' to enter a custom Canvas URL."
        )
    with col2:
        api_token = st.text_input(
            "API Token",
            type="password",
            help="Canvas → Account → Settings → New Access Token"
        )

    # Resolve canvas_url — show manual input only when "Other" is selected
    if UNIVERSITIES[selected_uni] is None:
        raw = st.text_input(
            "Canvas URL",
            placeholder="your-institution.instructure.com",
            help="Enter your institution's Canvas domain"
        )
        canvas_url = f"https://{raw}" if raw and not raw.startswith("http") else raw
    else:
        canvas_url = UNIVERSITIES[selected_uni]

    if st.button("Fetch My Courses", type="primary", disabled=not (canvas_url and api_token)):
        with st.spinner("Connecting to Canvas…"):
            st.session_state.courses = get_courses(canvas_url, api_token)
            if st.session_state.courses:
                st.success(f"Found {len(st.session_state.courses)} active course(s).")
            else:
                st.error("No courses found. Verify your Canvas URL and API token.")

    # ── Step 2 · Course selection ──────────────────────────
    if st.session_state.courses:
        step_header(2, "Select Course")

        course_options = {f"{c['name']}": c['id'] for c in st.session_state.courses}
        selected_course = st.selectbox("Course", options=list(course_options.keys()))
        course_id = course_options[selected_course]

        if st.button("Load Modules", disabled=not course_id):
            with st.spinner("Loading modules…"):
                st.session_state.modules = get_modules(canvas_url, api_token, course_id)
                if st.session_state.modules:
                    st.success(f"Found {len(st.session_state.modules)} module(s).")
                else:
                    st.warning("No modules found in this course.")

    # ── Step 3 · Module selection ──────────────────────────
    if st.session_state.modules:
        step_header(3, "Choose Modules", "Select which modules to include")

        select_all = st.checkbox("Include all modules", value=True)

        if select_all:
            selected_module_ids = [m['id'] for m in st.session_state.modules]
        else:
            selected_modules = st.multiselect(
                "Modules",
                options=[m['name'] for m in st.session_state.modules],
                default=[m['name'] for m in st.session_state.modules]
            )
            selected_module_ids = [
                m['id'] for m in st.session_state.modules
                if m['name'] in selected_modules
            ]

        n = len(selected_module_ids)
        st.markdown(
            f'<div class="count-badge">&#x25A0;&nbsp; {n} module{"s" if n != 1 else ""} selected</div>',
            unsafe_allow_html=True
        )

        # ── Step 4 · Download ──────────────────────────────
        step_header(4, "Download", "PPT, PPTX, PPTM & PDF files")

        if st.button("Download Files", type="primary", disabled=not selected_module_ids):
            progress_bar = st.progress(0)
            status_text = st.empty()

            try:
                zip_buffer, total_files = download_powerpoints(
                    canvas_url, api_token, course_id,
                    selected_module_ids, progress_bar, status_text
                )

                status_text.empty()
                progress_bar.empty()

                if total_files > 0:
                    st.success(f"Downloaded {total_files} file(s) successfully.")
                    st.download_button(
                        label="Save ZIP Archive",
                        data=zip_buffer,
                        file_name="canvas_files.zip",
                        mime="application/zip"
                    )
                else:
                    st.warning("No PowerPoint or PDF files found in the selected modules.")

            except Exception as e:
                st.error(f"Download failed: {e}")
                import traceback
                st.code(traceback.format_exc())

    # ── Footer ─────────────────────────────────────────────
    st.markdown("""
    <div class="tip-card">
        <strong>How to get your API token</strong><br>
        In Canvas, go to <strong>Account → Settings</strong>, scroll to
        <strong>Approved Integrations</strong>, and click
        <code>+ New Access Token</code>. Copy the token and paste it above.
    </div>
    """, unsafe_allow_html=True)


if __name__ == "__main__":
    main()

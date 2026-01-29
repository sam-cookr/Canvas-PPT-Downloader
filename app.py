#!/usr/bin/env python3
"""
Canvas PowerPoint Downloader - Streamlit Web App
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


class FileURLExtractor(HTMLParser):
    """Extract file URLs from HTML content"""

    def __init__(self):
        super().__init__()
        self.file_urls = []

    def handle_starttag(self, tag, attrs):
        if tag == 'a':
            for attr, value in attrs:
                if attr == 'href' and value:
                    if '/files/' in value or value.endswith(('.ppt', '.pptx', '.pptm')):
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


def is_powerpoint(filename):
    """Check if file is a PowerPoint"""
    return filename.lower().endswith(('.ppt', '.pptx', '.pptm'))


def sanitize_filename(filename):
    """Remove invalid characters from filename"""
    invalid_chars = '<>:"/\\|?*'
    for char in invalid_chars:
        filename = filename.replace(char, '_')
    return filename


def download_powerpoints(canvas_url, api_token, course_id, selected_module_ids, progress_bar, status_text):
    """Download PowerPoints from selected modules"""
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

            status_text.text(f"Processing module: {module_name}")

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

                    if is_powerpoint(filename):
                        file_url = item.get('url')
                        if file_url:
                            file_response = requests.get(file_url, headers=headers)
                            file_data = file_response.json()
                            download_url = file_data.get('url')

                            if download_url and download_url not in downloaded_files:
                                status_text.text(f"Downloading: {filename}")
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

                                    if is_powerpoint(filename):
                                        download_url = file_info.get('url')

                                        if download_url and download_url not in downloaded_files:
                                            status_text.text(f"Downloading: {filename}")
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


def main():
    st.set_page_config(
        page_title="Canvas PowerPoint Downloader",
        page_icon="📚",
        layout="centered"
    )

    st.title("📚 Canvas PowerPoint Downloader")
    st.markdown("Download all PowerPoint files from your Canvas course modules")

    # Initialize session state
    if 'courses' not in st.session_state:
        st.session_state.courses = []
    if 'modules' not in st.session_state:
        st.session_state.modules = []

    # Configuration section
    st.header("⚙️ Configuration")

    col1, col2 = st.columns(2)

    with col1:
        canvas_url = st.text_input(
            "Canvas URL",
            placeholder="newcastle.instructure.com",
            help="Enter your institution's Canvas URL"
        )

    with col2:
        api_token = st.text_input(
            "API Token",
            type="password",
            help="Your Canvas API token"
        )

    # Add https:// if not present
    if canvas_url and not canvas_url.startswith('http'):
        canvas_url = f"https://{canvas_url}"

    # Fetch courses button
    if st.button("🔍 Fetch My Courses", type="primary", disabled=not (canvas_url and api_token)):
        with st.spinner("Fetching courses..."):
            st.session_state.courses = get_courses(canvas_url, api_token)
            if st.session_state.courses:
                st.success(f"Found {len(st.session_state.courses)} courses!")
            else:
                st.error("No courses found. Check your Canvas URL and API token.")

    # Course selection
    if st.session_state.courses:
        st.header("📖 Select Course")

        course_options = {f"{c['name']} (ID: {c['id']})": c['id'] for c in st.session_state.courses}
        selected_course = st.selectbox(
            "Choose a course",
            options=list(course_options.keys())
        )

        course_id = course_options[selected_course]

        # Fetch modules button
        if st.button("📂 Load Modules", disabled=not course_id):
            with st.spinner("Loading modules..."):
                st.session_state.modules = get_modules(canvas_url, api_token, course_id)
                if st.session_state.modules:
                    st.success(f"Found {len(st.session_state.modules)} modules!")
                else:
                    st.warning("No modules found in this course.")

        # Module selection
        if st.session_state.modules:
            st.header("📋 Select Modules")

            select_all = st.checkbox("Select all modules", value=True)

            if select_all:
                selected_module_ids = [m['id'] for m in st.session_state.modules]
            else:
                selected_modules = st.multiselect(
                    "Choose modules to download",
                    options=[m['name'] for m in st.session_state.modules],
                    default=[m['name'] for m in st.session_state.modules]
                )
                selected_module_ids = [
                    m['id'] for m in st.session_state.modules
                    if m['name'] in selected_modules
                ]

            st.info(f"Selected {len(selected_module_ids)} module(s)")

            # Download button
            st.header("⬇️ Download")

            if st.button("📥 Download PowerPoints", type="primary", disabled=not selected_module_ids):
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
                        st.success(f"✅ Successfully downloaded {total_files} PowerPoint file(s)!")

                        st.download_button(
                            label="💾 Download ZIP File",
                            data=zip_buffer,
                            file_name="canvas_powerpoints.zip",
                            mime="application/zip"
                        )
                    else:
                        st.warning("No PowerPoint files found in the selected modules.")

                except Exception as e:
                    st.error(f"Error: {e}")
                    import traceback
                    st.code(traceback.format_exc())

    # Footer
    st.divider()
    st.markdown(
        """
        <div style='text-align: center; color: #666; font-size: 0.9em;'>
            <p>💡 Tip: You can find your API token in Canvas under Account → Settings → New Access Token</p>
        </div>
        """,
        unsafe_allow_html=True
    )


if __name__ == "__main__":
    main()

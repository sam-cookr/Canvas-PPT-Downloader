#!/usr/bin/env python3
"""
Canvas PowerPoint Downloader
Downloads all PowerPoint files from modules in a Canvas course
Handles both direct file links and files embedded in Pages
"""

import requests
import os
import re
from pathlib import Path
from urllib.parse import urljoin, urlparse, parse_qs
from html.parser import HTMLParser


class FileURLExtractor(HTMLParser):
    """Extract file URLs from HTML content"""

    def __init__(self):
        super().__init__()
        self.file_urls = []

    def handle_starttag(self, tag, attrs):
        if tag == 'a':
            for attr, value in attrs:
                if attr == 'href' and value:
                    # Check if it's a file URL
                    if '/files/' in value or value.endswith(('.ppt', '.pptx', '.pptm')):
                        self.file_urls.append(value)


def verify_course_access(canvas_url, headers, course_id):
    """Verify that the course exists and user has access"""
    url = f"{canvas_url}/api/v1/courses/{course_id}"
    try:
        response = requests.get(url, headers=headers)
        response.raise_for_status()
        course = response.json()
        print(f"✅ Found course: {course.get('name', 'Unknown')}")
        return True
    except requests.exceptions.HTTPError as e:
        if e.response.status_code == 401:
            print("❌ Authentication failed. Please check your API token.")
        elif e.response.status_code == 404:
            print(f"❌ Course with ID '{course_id}' not found or you don't have access.")
            print("   Try listing your courses first to find the correct ID.")
        else:
            print(f"❌ HTTP Error: {e}")
        return False
    except Exception as e:
        print(f"❌ Error verifying course: {e}")
        return False


def list_courses(canvas_url, headers):
    """List all available courses for the user"""
    url = f"{canvas_url}/api/v1/courses"
    params = {'per_page': 100, 'enrollment_state': 'active'}

    try:
        response = requests.get(url, headers=headers, params=params)
        response.raise_for_status()
        courses = response.json()

        if courses:
            print("\n📚 Your available courses:")
            for course in courses:
                if 'name' in course and 'id' in course:
                    print(f"  {course['id']}: {course['name']}")
            print()
        else:
            print("⚠️  No courses found")
    except Exception as e:
        print(f"⚠️  Could not fetch courses: {e}")


def get_modules(canvas_url, headers, course_id):
    """Fetch all modules in the course"""
    url = f"{canvas_url}/api/v1/courses/{course_id}/modules"
    params = {'per_page': 100}

    modules = []
    while url:
        response = requests.get(url, headers=headers, params=params)
        response.raise_for_status()
        modules.extend(response.json())

        # Check for pagination
        if 'next' in response.links:
            url = response.links['next']['url']
            params = {}  # Params are in the URL now
        else:
            url = None

    return modules


def get_module_items(canvas_url, headers, course_id, module_id):
    """Fetch all items in a module"""
    url = f"{canvas_url}/api/v1/courses/{course_id}/modules/{module_id}/items"
    params = {'per_page': 100}

    items = []
    while url:
        response = requests.get(url, headers=headers, params=params)
        response.raise_for_status()
        items.extend(response.json())

        # Check for pagination
        if 'next' in response.links:
            url = response.links['next']['url']
            params = {}
        else:
            url = None

    return items


def get_page_content(headers, page_url):
    """Fetch the content of a Canvas page"""
    try:
        response = requests.get(page_url, headers=headers)
        response.raise_for_status()
        page_data = response.json()
        return page_data.get('body', '')
    except Exception as e:
        print(f"    ⚠️  Could not fetch page content: {e}")
        return ''


def extract_file_id_from_url(url):
    """Extract Canvas file ID from various URL formats"""
    # Pattern 1: /courses/XXX/files/YYY
    match = re.search(r'/files/(\d+)', url)
    if match:
        return match.group(1)

    # Pattern 2: download?download_frd=1 with preview or other params
    if 'download_frd' in url or 'wrap' in url:
        match = re.search(r'/files/(\d+)', url)
        if match:
            return match.group(1)

    return None


def get_file_info(canvas_url, headers, course_id, file_id):
    """Get file information from Canvas API"""
    try:
        url = f"{canvas_url}/api/v1/courses/{course_id}/files/{file_id}"
        response = requests.get(url, headers=headers)
        response.raise_for_status()
        return response.json()
    except Exception as e:
        print(f"    ⚠️  Could not get file info for ID {file_id}: {e}")
        return None


def download_file(headers, url, filepath):
    """Download a file from URL to filepath"""
    response = requests.get(url, headers=headers, stream=True)
    response.raise_for_status()

    with open(filepath, 'wb') as f:
        for chunk in response.iter_content(chunk_size=8192):
            f.write(chunk)


def is_powerpoint(filename):
    """Check if file is a PowerPoint"""
    return filename.lower().endswith(('.ppt', '.pptx', '.pptm'))


def sanitize_filename(filename):
    """Remove invalid characters from filename"""
    invalid_chars = '<>:"/\\|?*'
    for char in invalid_chars:
        filename = filename.replace(char, '_')
    return filename


def main():
    # Get Canvas URL
    canvas_url = input("Enter your Canvas URL (e.g., newcastle.instructure.com): ").strip()
    if not canvas_url.startswith('http'):
        canvas_url = f"https://{canvas_url}"

    # Get API token
    api_token = input("Enter your Canvas API token: ").strip()

    # Validate API token
    if not api_token or len(api_token) < 10:
        print("❌ Error: API token appears to be invalid (too short)")
        return

    # Headers for API requests
    headers = {
        'Authorization': f'Bearer {api_token}'
    }

    # Offer to list courses
    list_first = input("Would you like to see a list of your courses first? (y/n): ").strip().lower()
    if list_first == 'y':
        list_courses(canvas_url, headers)

    # Get course ID
    course_id = input("Enter your course ID (should be a number, e.g., 12345): ").strip()

    # Validate course ID - must be numeric
    if not course_id.isdigit():
        print(f"❌ Error: Course ID must be a number, not '{course_id}'")
        print("   Tip: Find your course ID in the URL: https://your-canvas.com/courses/12345")
        print("   The course ID would be '12345'")
        return

    # Output directory
    output_dir = Path("canvas_powerpoints")
    output_dir.mkdir(exist_ok=True)

    print(f"\n🔍 Verifying access to course {course_id}...")

    # First verify the course exists and is accessible
    if not verify_course_access(canvas_url, headers, course_id):
        return

    print(f"\n🔍 Fetching modules from course {course_id}...")

    try:
        modules = get_modules(canvas_url, headers, course_id)
        print(f"✅ Found {len(modules)} modules")

        total_ppts = 0
        downloaded_files = set()  # Track to avoid duplicates

        for module in modules:
            module_name = module['name']
            module_id = module['id']

            print(f"\n📂 Processing module: {module_name}")

            # Create subfolder for this module
            module_folder = output_dir / sanitize_filename(module_name)
            module_folder.mkdir(exist_ok=True)

            # Get all items in the module
            items = get_module_items(canvas_url, headers, course_id, module_id)

            for item in items:
                item_type = item.get('type')
                item_title = item.get('title', 'unknown')

                # Handle direct File items
                if item_type == 'File':
                    filename = item_title

                    if is_powerpoint(filename):
                        file_url = item.get('url')
                        if file_url:
                            print(f"  📥 Found direct file: {filename}")
                            file_response = requests.get(file_url, headers=headers)
                            file_data = file_response.json()
                            download_url = file_data.get('url')

                            if download_url and download_url not in downloaded_files:
                                filepath = module_folder / sanitize_filename(filename)
                                download_file(headers, download_url, filepath)
                                downloaded_files.add(download_url)
                                total_ppts += 1
                                print(f"  ✅ Downloaded: {filename}")

                # Handle Page items (this is where your PPTs are!)
                elif item_type == 'Page':
                    print(f"  📄 Checking page: {item_title}")
                    page_url = item.get('url')

                    if page_url:
                        # Get the page content
                        page_html = get_page_content(headers, page_url)

                        # Extract file URLs from the HTML
                        parser = FileURLExtractor()
                        parser.feed(page_html)

                        # Process each file URL found
                        for file_url in parser.file_urls:
                            # Extract file ID
                            file_id = extract_file_id_from_url(file_url)

                            if file_id:
                                # Get file info from API
                                file_info = get_file_info(canvas_url, headers, course_id, file_id)

                                if file_info:
                                    filename = file_info.get('display_name', file_info.get('filename', 'unknown'))

                                    if is_powerpoint(filename):
                                        download_url = file_info.get('url')

                                        if download_url and download_url not in downloaded_files:
                                            print(f"    📥 Found in page: {filename}")
                                            filepath = module_folder / sanitize_filename(filename)
                                            download_file(headers, download_url, filepath)
                                            downloaded_files.add(download_url)
                                            total_ppts += 1
                                            print(f"    ✅ Downloaded: {filename}")

        print(f"\n🎉 Complete! Downloaded {total_ppts} PowerPoint files to '{output_dir}'")

    except requests.exceptions.HTTPError as e:
        if e.response.status_code == 401:
            print("\n❌ Authentication failed. Please check your API token.")
        elif e.response.status_code == 404:
            print("\n❌ Course not found. Please check your course ID.")
        else:
            print(f"\n❌ HTTP Error: {e}")
    except Exception as e:
        print(f"\n❌ Error: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()

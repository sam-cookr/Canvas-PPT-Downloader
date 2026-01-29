#!/usr/bin/env python3
"""
Flatten Canvas PowerPoint Directory (with options)
Advanced version with multiple flattening strategies
"""

from pathlib import Path
import shutil

# Configuration
SOURCE_DIR = Path("canvas_powerpoints")
OUTPUT_DIR = Path("all_powerpoints")


def is_powerpoint(filename):
    """Check if file is a PowerPoint"""
    return filename.lower().endswith(('.ppt', '.pptx', '.pptm'))


def flatten_simple():
    """Copy all files, rename duplicates only"""
    OUTPUT_DIR.mkdir(exist_ok=True)
    print(f"📁 Simple flatten: '{SOURCE_DIR}' → '{OUTPUT_DIR}'...\n")

    copied_count = 0

    for ppt_file in SOURCE_DIR.rglob("*"):
        if ppt_file.is_file() and is_powerpoint(ppt_file.name):
            module_name = ppt_file.parent.name
            dest_path = OUTPUT_DIR / ppt_file.name

            # Handle duplicates
            if dest_path.exists():
                counter = 2
                stem = ppt_file.stem
                suffix = ppt_file.suffix
                while dest_path.exists():
                    new_name = f"{stem} ({counter}){suffix}"
                    dest_path = OUTPUT_DIR / new_name
                    counter += 1
                print(f"  ⚠️  Renamed duplicate: {dest_path.name}")

            shutil.copy2(ppt_file, dest_path)
            copied_count += 1
            print(f"  ✅ {dest_path.name}")

    return copied_count


def flatten_with_prefixes():
    """Add module name prefix to ALL files"""
    OUTPUT_DIR.mkdir(exist_ok=True)
    print(f"📁 Flatten with module prefixes: '{SOURCE_DIR}' → '{OUTPUT_DIR}'...\n")

    copied_count = 0

    for ppt_file in SOURCE_DIR.rglob("*"):
        if ppt_file.is_file() and is_powerpoint(ppt_file.name):
            module_name = ppt_file.parent.name

            # Always add module prefix
            stem = ppt_file.stem
            suffix = ppt_file.suffix
            new_name = f"{module_name} - {stem}{suffix}"
            dest_path = OUTPUT_DIR / new_name

            # Handle rare case of still having duplicates
            counter = 2
            while dest_path.exists():
                new_name = f"{module_name} - {stem} ({counter}){suffix}"
                dest_path = OUTPUT_DIR / new_name
                counter += 1

            shutil.copy2(ppt_file, dest_path)
            copied_count += 1
            print(f"  ✅ {dest_path.name}")

    return copied_count


def flatten_by_number():
    """Rename all files with sequential numbers"""
    OUTPUT_DIR.mkdir(exist_ok=True)
    print(f"📁 Flatten with numbering: '{SOURCE_DIR}' → '{OUTPUT_DIR}'...\n")

    # Collect all files first
    all_files = []
    for ppt_file in SOURCE_DIR.rglob("*"):
        if ppt_file.is_file() and is_powerpoint(ppt_file.name):
            all_files.append(ppt_file)

    # Sort by module name and filename
    all_files.sort(key=lambda x: (x.parent.name, x.name))

    copied_count = 0
    for idx, ppt_file in enumerate(all_files, 1):
        module_name = ppt_file.parent.name
        suffix = ppt_file.suffix

        # Create numbered filename
        new_name = f"{idx:02d} - {module_name} - {ppt_file.stem}{suffix}"
        dest_path = OUTPUT_DIR / new_name

        shutil.copy2(ppt_file, dest_path)
        copied_count += 1
        print(f"  ✅ {new_name}")

    return copied_count


def main():
    if not SOURCE_DIR.exists():
        print(f"❌ Source directory '{SOURCE_DIR}' not found!")
        print(f"   Make sure you've run the download script first.")
        return

    print("🔧 Choose flattening method:\n")
    print("1. Simple (keep original names, rename duplicates only)")
    print("2. With module prefixes (add module name to all files)")
    print("3. Numbered (sequential numbering with module names)")
    print()

    choice = input("Enter choice (1-3): ").strip()
    print()

    if choice == "1":
        copied = flatten_simple()
    elif choice == "2":
        copied = flatten_with_prefixes()
    elif choice == "3":
        copied = flatten_by_number()
    else:
        print("❌ Invalid choice!")
        return

    print(f"\n🎉 Complete! Copied {copied} PowerPoint files to '{OUTPUT_DIR}'")


if __name__ == "__main__":
    main()
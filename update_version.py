#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os
import sys
import re
import platform

# Force UTF-8 encoding in Windows environment
if platform.system() == 'Windows':
    # Set console code page to UTF-8
    os.system('chcp 65001 > nul')
    # Force stdout and stderr to use UTF-8
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')
    if hasattr(sys.stderr, 'reconfigure'):
        sys.stderr.reconfigure(encoding='utf-8')

# Print system information for debugging
print(f"System: {platform.system()}")
print(f"Python version: {platform.python_version()}")
print(f"Default encoding: {sys.getdefaultencoding()}")

# Read version from MC_Recon_UI.py
print("Reading version from MC_Recon_UI.py...")
try:
    with open('MC_Recon_UI.py', 'r', encoding='utf-8') as f:
        content = f.read()
        version_match = re.search(r"VERSION = '([\d\.]+)'\s*", content)
        if version_match:
            version = version_match.group(1)
            print(f"Found version: {version}")
        else:
            print("Error: Could not find version in MC_Recon_UI.py")
            sys.exit(1)
except Exception as e:
    print(f"Error reading MC_Recon_UI.py: {e}")
    sys.exit(1)

# Parse version
version_parts = version.split('.')
if len(version_parts) != 3:
    print(f"Error: Invalid version format: {version}")
    sys.exit(1)

major, minor, patch = map(int, version_parts)

# Increment patch version
patch += 1
new_version = f"{major}.{minor}.{patch}"
print(f"Incrementing patch version: {version} -> {new_version}")

# Update version in MC_Recon_UI.py
print("Updating version in MC_Recon_UI.py...")
try:
    new_content = re.sub(r"VERSION = '[\d\.]+'\s*", f"VERSION = '{new_version}'\n", content)
    with open('MC_Recon_UI.py', 'w', encoding='utf-8') as f:
        f.write(new_content)
    print(f"Updated version in MC_Recon_UI.py to {new_version}")
except Exception as e:
    print(f"Error updating MC_Recon_UI.py: {e}")
    sys.exit(1)

# Update version in file_version_info.txt
print("Updating version in file_version_info.txt...")
try:
    with open('file_version_info.txt', 'r', encoding='utf-8') as f:
        content = f.read()
    
    # Update filevers and prodvers
    content = re.sub(r"filevers=\([\d, ]+\)", f"filevers=({major}, {minor}, {patch}, 0)", content)
    content = re.sub(r"prodvers=\([\d, ]+\)", f"prodvers=({major}, {minor}, {patch}, 0)", content)
    
    # Update FileVersion and ProductVersion
    content = re.sub(
        r"StringStruct\('FileVersion', '[\d\.]+'\)",
        f"StringStruct('FileVersion', '{new_version}')",
        content
    )
    content = re.sub(
        r"StringStruct\('ProductVersion', '[\d\.]+'\)",
        f"StringStruct('ProductVersion', '{new_version}')",
        content
    )
    
    with open('file_version_info.txt', 'w', encoding='utf-8') as f:
        f.write(content)
    print(f"Updated version in file_version_info.txt to {new_version}")
except Exception as e:
    print(f"Error updating file_version_info.txt: {e}")
    sys.exit(1)

print("Version update completed successfully!")

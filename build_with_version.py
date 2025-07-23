#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os
import sys
import shutil
import zipfile
import platform
import subprocess
from pathlib import Path

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

# Update version
print("\n" + "=" * 50)
print("Step 1: Updating version number")
print("=" * 50)

try:
    subprocess.run([sys.executable, 'update_version.py'], check=True)
    print("Version updated successfully!")
except subprocess.CalledProcessError as e:
    print(f"Error updating version: {e}")
    sys.exit(1)

# Get current version
with open('MC_Recon_UI.py', 'r', encoding='utf-8') as f:
    content = f.read()
    import re
    version_match = re.search(r"VERSION = '([\d\.]+)'\s*", content)
    if version_match:
        current_version = version_match.group(1)
        print(f"Current version: {current_version}")
    else:
        print("Error: Could not find version in MC_Recon_UI.py")
        sys.exit(1)

# Compile resources
print("\n" + "=" * 50)
print("Step 2: Compiling resource files")
print("=" * 50)

try:
    subprocess.run(['pyrcc5', 'resources.qrc', '-o', 'resources.py'], check=True)
    print("Resource files compiled successfully!")
except subprocess.CalledProcessError as e:
    print(f"Error compiling resources: {e}")
    sys.exit(1)

# Build with PyInstaller
print("\n" + "=" * 50)
print("Step 3: Building with PyInstaller")
print("=" * 50)

try:
    subprocess.run(['pyinstaller', f'MC对账明细工具.spec'], check=True)
    print("PyInstaller build completed successfully!")
except subprocess.CalledProcessError as e:
    print(f"Error building with PyInstaller: {e}")
    sys.exit(1)

# Create zip file
print("\n" + "=" * 50)
print("Step 4: Creating ZIP package")
print("=" * 50)

exe_name = f"MC对账明细工具_v{current_version}.exe"
exe_path = os.path.join('dist', exe_name)

if not os.path.exists(exe_path):
    print(f"Error: Executable file not found at {exe_path}")
    # Try to find the actual file
    dist_files = os.listdir('dist')
    print(f"Files in dist directory: {dist_files}")
    sys.exit(1)

zip_name = f"MC对账明细工具_v{current_version}.zip"
zip_path = os.path.join('dist', zip_name)

try:
    with zipfile.ZipFile(zip_path, 'w') as zipf:
        # Add executable
        zipf.write(exe_path, os.path.basename(exe_path))
        print(f"Added {exe_name} to ZIP")
        
        # Add config.ini if it exists
        if os.path.exists('config.ini'):
            zipf.write('config.ini', 'config.ini')
            print("Added config.ini to ZIP")
        else:
            print("Warning: config.ini not found, skipping")
            
    print(f"ZIP package created successfully: {zip_path}")
    print(f"ZIP file size: {os.path.getsize(zip_path) / (1024*1024):.2f} MB")
except Exception as e:
    print(f"Error creating ZIP package: {e}")
    sys.exit(1)

print("\n" + "=" * 50)
print("Build process completed successfully!")
print("=" * 50)

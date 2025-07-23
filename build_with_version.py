#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os
import sys
import shutil
import zipfile
import platform
import subprocess
import re
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
try:
    with open('MC_Recon_UI.py', 'r', encoding='utf-8') as f:
        content = f.read()
        print(f"File size: {len(content)} bytes")
        
        # Try multiple patterns to find version
        patterns = [
            r"VERSION\s*=\s*'([\d\.]+)'\s*",  # VERSION = '1.1.16'
            r"VERSION\s*=\s*'([\d\.]+)'",      # VERSION = '1.1.16' (no trailing whitespace)
            r"VERSION\s*=\s*\"([\d\.]+)\"\s*", # VERSION = "1.1.16"
            r"VERSION\s*=\s*\"([\d\.]+)\"",    # VERSION = "1.1.16" (no trailing whitespace)
            r"self\.version\s*=\s*'([\d\.]+)'\s*", # self.version = '1.1.16'
            r"self\.version\s*=\s*\"([\d\.]+)\"\s*", # self.version = "1.1.16"
            r"__version__\s*=\s*'([\d\.]+)'\s*", # __version__ = '1.1.16'
            r"__version__\s*=\s*\"([\d\.]+)\"\s*" # __version__ = "1.1.16"
        ]
        
        current_version = None
        for i, pattern in enumerate(patterns):
            version_match = re.search(pattern, content)
            if version_match:
                current_version = version_match.group(1)
                print(f"Found version: {current_version} using pattern {i+1}")
                break
        
        if not current_version:
            # Try to find any version-like pattern
            backup_pattern = r"['\"]([\d]+\.[\d]+\.[\d]+)['\"]"  # '1.1.16' or "1.1.16"
            version_match = re.search(backup_pattern, content)
            if version_match:
                current_version = version_match.group(1)
                print(f"Found version with backup pattern: {current_version}")
            else:
                # Print file content sample for debugging
                print("Error: Could not find version in MC_Recon_UI.py")
                print("File content sample:")
                lines = content.split('\n')
                
                # Try to find lines with version-like strings
                version_lines = []
                for i, line in enumerate(lines):
                    if re.search(r"['\"][\d]+\.[\d]+\.[\d]+['\"]|version|VERSION", line, re.IGNORECASE):
                        version_lines.append((i, line))
                
                if version_lines:
                    print("Found lines that might contain version information:")
                    for i, line in version_lines[:10]:  # Show at most 10 lines
                        print(f"Line {i+1}: {line}")
                else:
                    # Show a sample of lines if no version-like lines found
                    for i in range(max(0, 580), min(585, len(lines))):
                        print(f"Line {i+1}: {lines[i]}")
                
                sys.exit(1)
                
        # If we get here, we found a version
        print(f"Using version: {current_version}")
        
except Exception as e:
    print(f"Error reading MC_Recon_UI.py: {e}")
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
    
    # Try to find any .exe file
    exe_files = [f for f in dist_files if f.endswith('.exe')]
    if exe_files:
        print(f"Found executable files: {exe_files}")
        exe_path = os.path.join('dist', exe_files[0])
        print(f"Using: {exe_path}")
    else:
        sys.exit(1)

zip_name = f"MC对账明细工具_v{current_version}.zip"
zip_path = os.path.join('dist', zip_name)

try:
    with zipfile.ZipFile(zip_path, 'w') as zipf:
        # Add executable
        zipf.write(exe_path, os.path.basename(exe_path))
        print(f"Added {os.path.basename(exe_path)} to ZIP")
        
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

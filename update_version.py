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
        
        version = None
        for i, pattern in enumerate(patterns):
            version_match = re.search(pattern, content)
            if version_match:
                version = version_match.group(1)
                print(f"Found version: {version} using pattern {i+1}")
                break
        
        if not version:
            # Try to find any version-like pattern
            backup_pattern = r"['\"]([\d]+\.[\d]+\.[\d]+)['\"]"  # '1.1.16' or "1.1.16"
            version_match = re.search(backup_pattern, content)
            if version_match:
                version = version_match.group(1)
                print(f"Found version with backup pattern: {version}")
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
        print(f"Using version: {version}")
        
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
    # Try to replace using the same pattern that matched
    updated = False
    for pattern in patterns:
        version_match = re.search(pattern, content)
        if version_match:
            # Use the same format for replacement
            old_version_str = version_match.group(0)
            new_version_str = old_version_str.replace(version, new_version)
            new_content = content.replace(old_version_str, new_version_str)
            
            # Check if replacement was successful
            if new_content != content:
                with open('MC_Recon_UI.py', 'w', encoding='utf-8') as f:
                    f.write(new_content)
                print(f"Updated version in MC_Recon_UI.py to {new_version}")
                updated = True
                break
    
    if not updated:
        # If no pattern matched or replacement failed, try a more generic approach
        print("Using generic replacement approach")
        new_content = re.sub(r"(['\"])[\d]+\.[\d]+\.[\d]+(['\"])", f"\1{new_version}\2", content, count=1)
        
        # Check if replacement was successful
        if new_content != content:
            with open('MC_Recon_UI.py', 'w', encoding='utf-8') as f:
                f.write(new_content)
            print(f"Updated version in MC_Recon_UI.py to {new_version} using generic replacement")
        else:
            print("Warning: Could not update version in MC_Recon_UI.py")
            # Try to add version if it doesn't exist
            if "VERSION = " not in content:
                # Add version definition at the beginning of the file
                new_content = f"VERSION = '{new_version}'\n\n" + content
                with open('MC_Recon_UI.py', 'w', encoding='utf-8') as f:
                    f.write(new_content)
                print(f"Added VERSION = '{new_version}' to the beginning of MC_Recon_UI.py")
        
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

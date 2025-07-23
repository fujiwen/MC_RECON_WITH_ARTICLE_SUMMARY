import re
import os
import sys
import platform
from datetime import datetime

# 设置控制台编码为UTF-8
os.environ["PYTHONIOENCODING"] = "utf-8"

# 在Windows环境下设置控制台代码页为UTF-8
if platform.system() == "Windows":
    try:
        # 设置控制台代码页为UTF-8 (65001)
        os.system("chcp 65001 > nul")
        # 强制使用UTF-8输出
        if hasattr(sys.stdout, 'reconfigure'):
            sys.stdout.reconfigure(encoding='utf-8')
            sys.stderr.reconfigure(encoding='utf-8')
    except Exception as e:
        print(f"Warning: Failed to set console encoding: {e}")

print(f"System: {platform.system()}, Python: {platform.python_version()}, Encoding: {sys.getdefaultencoding()}")


def update_version():
    # Read current version number
    version_pattern = re.compile(r"VERSION\s*=\s*['\"]([0-9]+)\.([0-9]+)\.([0-9]+)['\"]")
    
    mc_recon_path = 'MC_Recon_UI.py'
    version_file_path = 'file_version_info.txt'
    
    # Read MC_Recon_UI.py file content
    with open(mc_recon_path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # Find version number
    match = version_pattern.search(content)
    if not match:
        print("Could not find version number in MC_Recon_UI.py")
        return False
    
    # Parse version number
    major, minor, patch = map(int, match.groups())
    
    # Increment patch version
    patch += 1
    new_version = f"{major}.{minor}.{patch}"
    
    # Update version in MC_Recon_UI.py
    new_content = version_pattern.sub(f"VERSION = '{major}.{minor}.{patch}'", content)
    with open(mc_recon_path, 'w', encoding='utf-8') as f:
        f.write(new_content)
    
    # Update version in file_version_info.txt
    if os.path.exists(version_file_path):
        with open(version_file_path, 'r', encoding='utf-8') as f:
            version_content = f.read()
        
        # Update filevers and prodvers
        filevers_pattern = re.compile(r"filevers=\(([0-9]+),\s*([0-9]+),\s*([0-9]+),\s*([0-9]+)\)")
        version_content = filevers_pattern.sub(f"filevers=({major}, {minor}, {patch}, 0)", version_content)
        
        prodvers_pattern = re.compile(r"prodvers=\(([0-9]+),\s*([0-9]+),\s*([0-9]+),\s*([0-9]+)\)")
        version_content = prodvers_pattern.sub(f"prodvers=({major}, {minor}, {patch}, 0)", version_content)
        
        # Update FileVersion and ProductVersion
        file_version_pattern = re.compile(r"StringStruct\(u'FileVersion',\s*u'[0-9\.]+'\)")
        version_content = file_version_pattern.sub(f"StringStruct(u'FileVersion', u'{new_version}')", version_content)
        
        product_version_pattern = re.compile(r"StringStruct\(u'ProductVersion',\s*u'[0-9\.]+'\)")
        version_content = product_version_pattern.sub(f"StringStruct(u'ProductVersion', u'{new_version}')", version_content)
        
        with open(version_file_path, 'w', encoding='utf-8') as f:
            f.write(version_content)
    
    print(f"Version updated to: {new_version}")
    return True

if __name__ == "__main__":
    update_version()
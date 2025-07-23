import os
import re
import subprocess
import sys
import zipfile
import platform
from datetime import datetime

# Set console encoding to UTF-8
os.environ["PYTHONIOENCODING"] = "utf-8"

# Set console code page to UTF-8 in Windows environment
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


# Print header
print("="*50)
print("      MC Reconciliation Tool - Auto Packaging Script")
print("="*50)

# Update version number
print("[1/5] Updating version number...")
result = subprocess.run([sys.executable, "update_version.py"], capture_output=True, text=True, encoding="utf-8")
print(result.stdout)

# Get current version number
with open("MC_Recon_UI.py", "r", encoding="utf-8") as f:
    content = f.read()
    version_match = re.search(r"VERSION = '([\d\.]+)'", content)
    current_version = version_match.group(1) if version_match else "unknown"

# Compile resource file
print("[2/5] Compiling resource file...")
resource_result = subprocess.run(["pyrcc5", "-o", "resources.py", "resources.qrc"], capture_output=True, text=True, encoding="utf-8")
if resource_result.returncode != 0:
    print("Resource file compilation failed!")
    print("Error message:")
    print(resource_result.stderr)
    sys.exit(1)

# Package application
print(f"[3/5] Packaging application v{current_version}...")
result = subprocess.run(["pyinstaller", "mc_recon_tool.spec", "--clean"], capture_output=True, text=True, encoding="utf-8")

# Check packaging result
exe_path = os.path.join("dist", f"MC_Recon_Tool_v{current_version}.exe")
if os.path.exists(exe_path):
    print("[4/5] Packaging completed!")
    
    # Get file information
    file_time = datetime.fromtimestamp(os.path.getmtime(exe_path))
    file_size = os.path.getsize(exe_path) / (1024 * 1024)  # Convert to MB
    
    print("\nApplication information:")
    print(f"  File name: {exe_path}")
    print(f"  File size: {file_size:.2f} MB")
    print(f"  Creation time: {file_time.strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"  Version: {current_version}")
    
    # Create zip package
    print("[5/5] Creating zip package...")
    zip_filename = f"MC_Recon_Tool_v{current_version}.zip"
    zip_path = os.path.join("dist", zip_filename)
    
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        # Add newly packaged exe file
        zipf.write(exe_path, os.path.basename(exe_path))
        
        # Add configuration file
        config_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'config.ini')
        if os.path.exists(config_path):
            zipf.write(config_path, 'config.ini')
            print(f"Added configuration file: config.ini")
        else:
            print(f"Warning: Configuration file does not exist: {config_path}")
    
    # Get zip package information
    zip_size = os.path.getsize(zip_path) / (1024 * 1024)  # Convert to MB
    
    print("\nZip package information:")
    print(f"  File name: {zip_filename}")
    print(f"  File size: {zip_size:.2f} MB")
    
    # Show output directory
    print(f"\nPackage files located at: {os.path.abspath('dist')}")
    print(f"Build successful: {exe_path}")
else:
    print("[3/5] Packaging failed!")
    print("Error message:")
    print(result.stderr)

# In GitHub Actions environment, don't wait for input
if not os.environ.get('GITHUB_ACTIONS'):
    input("\nPress Enter to exit...")
import os
import re
import subprocess
import sys
import zipfile
import locale
from datetime import datetime

# 设置控制台编码为UTF-8
os.environ["PYTHONIOENCODING"] = "utf-8"

# 在Windows环境中设置控制台代码页为UTF-8
if sys.platform == "win32":
    try:
        # 设置控制台代码页为65001 (UTF-8)
        subprocess.run(["chcp", "65001"], shell=True, check=True)
        # 强制使用UTF-8输出
        sys.stdout.reconfigure(encoding='utf-8')
        sys.stderr.reconfigure(encoding='utf-8')
    except Exception as e:
        print(f"Warning: Failed to set console encoding: {e}")

# 打印标题
print("="*50)
print("      MC Recon Tool - Auto Build Script")
print("="*50)

# 更新版本号
print("[1/5] Updating version number...")
result = subprocess.run([sys.executable, "update_version.py"], capture_output=True, text=True, encoding="utf-8")
print(result.stdout)

# 获取当前版本号
with open("MC_Recon_UI.py", "r", encoding="utf-8") as f:
    content = f.read()
    version_match = re.search(r"VERSION = '([\d\.]+)'", content)
    current_version = version_match.group(1) if version_match else "unknown"

# 编译资源文件
print("[2/5] Compiling resource files...")
resource_result = subprocess.run(["pyrcc5", "-o", "resources.py", "resources.qrc"], capture_output=True, text=True, encoding="utf-8")
if resource_result.returncode != 0:
    print("Resource compilation failed!")
    print("Error message:")
    print(resource_result.stderr)
    sys.exit(1)

# 打包应用程序
print(f"[3/5] Packaging application v{current_version}...")
result = subprocess.run(["pyinstaller", "MC对账明细工具.spec", "--clean"], capture_output=True, text=True, encoding="utf-8")

# 检查打包结果
exe_path = os.path.join("dist", f"MC对账明细工具_v{current_version}.exe")
if os.path.exists(exe_path):
    print("[4/5] Packaging completed!")
    
    # 获取文件信息
    file_time = datetime.fromtimestamp(os.path.getmtime(exe_path))
    file_size = os.path.getsize(exe_path) / (1024 * 1024)  # 转换为MB
    
    print("\nApplication information:")
    print(f"  File name: {exe_path}")
    print(f"  File size: {file_size:.2f} MB")
    print(f"  Creation time: {file_time.strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"  Version: {current_version}")
    
    # 创建发布包
    print("\n[5/5] Creating release package...")
    release_dir = "release"
    if not os.path.exists(release_dir):
        os.makedirs(release_dir)
    
    # 创建zip文件
    zip_filename = os.path.join(release_dir, f"MC对账明细工具_v{current_version}.zip")
    with zipfile.ZipFile(zip_filename, 'w', zipfile.ZIP_DEFLATED) as zipf:
        zipf.write(exe_path, os.path.basename(exe_path))
    
    print(f"Release package created: {zip_filename}")
    print(f"File size: {os.path.getsize(zip_filename) / (1024 * 1024):.2f} MB")
    print("\nPackaging process completed!")
else:
    print("Packaging failed! Please check error message:")
    print(result.stderr)
    sys.exit(1)
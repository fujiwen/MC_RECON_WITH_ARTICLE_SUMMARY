import base64
import os
import hashlib
from PIL import Image

# 从GitHub API响应中获取的base64编码内容
github_icon_base64 = "AAABAAEAICAAAAEAIACoEAAAFgAAACgAAAAgAAAAQAAAAAEAIAAAAAAAABAAABMLAAATCwAAAAAAAAAAAADblhIA25YSANuWEgDblhIA25YSANuWEgDblhIA25YSANuWEgDblhIA25YSANuWEgDblhIA25YSANuWEgDblhIAAAAAAAAAAADalhIA2pYSG9uWEt7blhKi248NAduUEQDblhIA25YSANuWEprblhLk25YSH9uWEgAAAAAAAAAAANyUEQDblhIA25YSQ9uWEjnblhIA25YSLNuWEozblhKS25YSkduWEo/blhKP25YSkduWEpLblhKB25YSF9uWEgAAAAAAAAAAANqWEgDalhIb25YS3tuWEqLbjw0B25QRANuWEgDblhIA25YSmtuWEuTblhIf25YSAAAAAAAAAAAA3JQRANqYFADblhIA25YSANuWEgDblhIA25YSANuWEgDblhIA25YSANuWEgDblhIA25YSANuWEgDblhIA25YSAAAAAAAAAAAA2pYSANqWEhvblhLe25YSotuPDQHblBEA25YSANuWEgDblhKa25YS5NuWEh/blhIAAAAAAAAAAADZlRIA25YSANqWEgDblhIA25YSANuWEQDblhIA25YSANuWEgDblhIA25YSANuWEgDblhIA25YSAN2XEQDdlxEAAAAAAAAAAADalhIA2pYSG9uWEt7blhKi248NAduUEQDblhIA25YSANuWEprblhLk25YSHtuWEQDblREA25URANuVEQDblREA25URANuVEQDblREA25URANuVEQDblREA25URANuVEQDblREA25URANuVEQDblREA25URANuVEQDblREA25URANuVEQDblRIA2pYSGtuWEt7blhKi248NAduUEQDblhIA25YSANuWEprblhLq25YSMtuVEQvblREN25URDtuVEQ7blREN25URDtuVEQ7blREN25URDtuVEQ7blREN25URDtuVEQ7blREN25URDtuVEQ7blREN25URDtuVEQ7blREN25URDtuVEQzblhIt25YS5duWEqLbjw0B25QRANuWEgDblhIA25YSgNuWEv/blhLZ25YSyduWEsrblhLK25YSytuWEsrblhLK25YSytuWEsrblhLK25YSytuWEsrblhLK25YSytuWEsrblhLK25YSytuWEsrblhLK25YSytuWEsnblhLX25YS/9uWEojblxMA25URAOCYDwDblhIA25YSJNuWEqXblhLJ25YSytuWEsrblhLK25YSytuWEsrblhLK25YSytuWEsrblhLK25YSytuWEsrblhLK25YSytuWEsrblhLK25YSytuWEsrblhLK25YSytuWEsrblhLK25YSytuWEsnblhKp25YSKduWEgDakQ8A2ZYTANuWEgDblRIA25YSBduVEg3blREN25URDduVEQ3blREN25URDduVEQ3blREN25URDduVEQ3blREN25URDduVEQ3blREN25URDduVEQ3blREN25URDduVEQ3blREN25URDduVEQ3blREN25URDduWEgbblRIA2pUSANuXEgAAAAAA25cSANu4FADblhIA25URANuVEQDblREA25URANuVEQDblREA25URANuVEQDblREA25URANuVEQDblREA25URANuVEQDblREA25URANuVEQDblREA25URANuVEQDblREA25URANuVEQDblREA25YSAN+bEwDdmBIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA/wAA//8AAP//AAD//wAA//8AAP//AAD//wAA/wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAMCB8ADAgfAAwIHwAMCB8ADAgfAAwIHwAMAAMADAADAAwAAwAMAAMADAADAAwAAwAAAAAAAAAAAAAAAAAAAAAAAAAAAgAAAAf////8="

# 将base64编码的内容解码为二进制数据
github_icon_data = base64.b64decode(github_icon_base64)

# 保存GitHub图标到临时文件
github_icon_path = "github_favicon.ico"
with open(github_icon_path, "wb") as f:
    f.write(github_icon_data)

# 本地图标路径
local_icon_path = "favicon.ico"

# 计算两个文件的MD5哈希值进行比较
def get_file_md5(file_path):
    md5_hash = hashlib.md5()
    with open(file_path, "rb") as f:
        for byte_block in iter(lambda: f.read(4096), b""):
            md5_hash.update(byte_block)
    return md5_hash.hexdigest()

local_md5 = get_file_md5(local_icon_path)
github_md5 = get_file_md5(github_icon_path)

print(f"本地图标MD5: {local_md5}")
print(f"GitHub图标MD5: {github_md5}")
print(f"图标文件是否相同: {local_md5 == github_md5}")

# 使用PIL检查两个图标文件
try:
    local_img = Image.open(local_icon_path)
    github_img = Image.open(github_icon_path)
    
    print("\n本地图标信息:")
    print(f"Format: {local_img.format}, Size: {local_img.size}, Mode: {local_img.mode}")
    
    print("\nGitHub图标信息:")
    print(f"Format: {github_img.format}, Size: {github_img.size}, Mode: {github_img.mode}")
    
    # 检查图像内容是否相同
    if local_img.tobytes() == github_img.tobytes() and local_img.size == github_img.size and local_img.mode == github_img.mode:
        print("\n图像内容完全相同")
    else:
        print("\n图像内容不同")
        
except Exception as e:
    print(f"\n检查图像时出错: {e}")

# 清理临时文件
os.remove(github_icon_path)
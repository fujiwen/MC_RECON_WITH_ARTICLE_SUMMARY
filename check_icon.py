from PIL import Image

try:
    img = Image.open('favicon.ico')
    print('Icon file is valid and can be opened with PIL')
    print(f'Format: {img.format}, Size: {img.size}, Mode: {img.mode}')
except Exception as e:
    print(f'Error opening icon file: {e}')
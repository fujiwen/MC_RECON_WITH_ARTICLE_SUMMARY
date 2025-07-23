import base64

with open('new_favicon.ico', 'rb') as f:
    encoded = base64.b64encode(f.read()).decode('utf-8')
    print(encoded)
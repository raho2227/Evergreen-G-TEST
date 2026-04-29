from PIL import Image
import sys
try:
    img = Image.open('icon-192.png')
    img_192 = img.resize((192, 192), Image.Resampling.LANCZOS)
    img_192.save('icon-192.png', 'PNG')
    
    img_512 = img.resize((512, 512), Image.Resampling.LANCZOS)
    img_512.save('icon-512.png', 'PNG')
    print("Resized successfully in Python!")
except Exception as e:
    print(f"Error resizing: {e}")

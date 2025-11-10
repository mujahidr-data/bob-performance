#!/usr/bin/env python3
"""
Create a Bob-themed icon for the macOS app launcher
"""
import os
import sys

try:
    from PIL import Image, ImageDraw, ImageFont
    HAS_PIL = True
except ImportError:
    HAS_PIL = False
    print("⚠️  PIL/Pillow not installed. Creating icon using sips (macOS built-in)...")

def create_icon_with_pil(output_path):
    """Create icon using PIL/Pillow"""
    # Create a 512x512 image with Bob colors (blue/purple gradient)
    size = 512
    img = Image.new('RGBA', (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    
    # Draw gradient background (blue to purple)
    for i in range(size):
        r = int(102 + (118 - 102) * i / size)  # 102 to 118 (blue)
        g = int(126 + (76 - 126) * i / size)   # 126 to 76
        b = int(234 + (74 - 234) * i / size)  # 234 to 74 (purple)
        draw.rectangle([(0, i), (size, i+1)], fill=(r, g, b, 255))
    
    # Draw rounded rectangle background
    margin = 50
    draw.rounded_rectangle(
        [margin, margin, size-margin, size-margin],
        radius=80,
        fill=(255, 255, 255, 230)
    )
    
    # Draw "B" letter
    try:
        # Try to use a system font
        font_size = 300
        font = ImageFont.truetype("/System/Library/Fonts/Helvetica.ttc", font_size)
    except:
        try:
            font = ImageFont.truetype("/System/Library/Fonts/Arial.ttf", font_size)
        except:
            font = ImageFont.load_default()
    
    # Draw "B" in center
    bbox = draw.textbbox((0, 0), "B", font=font)
    text_width = bbox[2] - bbox[0]
    text_height = bbox[3] - bbox[1]
    x = (size - text_width) // 2
    y = (size - text_height) // 2 - 20
    
    # Draw "B" with gradient effect
    draw.text((x, y), "B", fill=(102, 126, 234, 255), font=font)
    
    # Save as PNG first
    png_path = output_path.replace('.icns', '.png')
    img.save(png_path, 'PNG')
    print(f"✅ Created icon PNG: {png_path}")
    
    return png_path

def create_icon_with_sips(png_path, icns_path):
    """Convert PNG to ICNS using macOS sips command"""
    if os.path.exists(icns_path):
        os.remove(icns_path)
    
    # Create iconset directory
    iconset_dir = icns_path.replace('.icns', '.iconset')
    os.makedirs(iconset_dir, exist_ok=True)
    
    # Generate all required sizes
    sizes = [16, 32, 64, 128, 256, 512, 1024]
    for size in sizes:
        # Regular size
        os.system(f'sips -z {size} {size} "{png_path}" --out "{iconset_dir}/icon_{size}x{size}.png" > /dev/null 2>&1')
        # Retina size (2x)
        if size < 1024:
            os.system(f'sips -z {size*2} {size*2} "{png_path}" --out "{iconset_dir}/icon_{size}x{size}@2x.png" > /dev/null 2>&1')
    
    # Convert iconset to icns
    os.system(f'iconutil -c icns "{iconset_dir}" -o "{icns_path}" > /dev/null 2>&1')
    
    # Clean up iconset directory
    import shutil
    shutil.rmtree(iconset_dir, ignore_errors=True)
    
    if os.path.exists(icns_path):
        print(f"✅ Created icon ICNS: {icns_path}")
        return True
    return False

def main():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    project_root = os.path.dirname(os.path.dirname(script_dir))
    icon_dir = os.path.join(project_root, "assets")
    os.makedirs(icon_dir, exist_ok=True)
    
    png_path = os.path.join(icon_dir, "bob_icon.png")
    icns_path = os.path.join(icon_dir, "bob_icon.icns")
    
    print("🎨 Creating Bob icon...")
    print("")
    
    if HAS_PIL:
        create_icon_with_pil(icns_path)
        png_path = icns_path.replace('.icns', '.png')
    else:
        # Create a simple icon using sips from a solid color
        # This is a fallback if PIL is not available
        print("Creating simple icon...")
        os.system(f'sips -s format png -z 512 512 --setColor "102 126 234" /dev/null --out "{png_path}" > /dev/null 2>&1')
    
    # Convert to ICNS
    if os.path.exists(png_path):
        create_icon_with_sips(png_path, icns_path)
        print("")
        print(f"✅ Icon created successfully!")
        print(f"   PNG: {png_path}")
        print(f"   ICNS: {icns_path}")
        return icns_path
    else:
        print("❌ Failed to create icon")
        return None

if __name__ == "__main__":
    main()


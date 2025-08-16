#!/usr/bin/env python3
"""
Logo and Favicon Generator for Excel Column Mapper
Creates professional-looking logo and favicon assets automatically.

Requirements: pip install Pillow
"""

from PIL import Image, ImageDraw, ImageFont
import os
from pathlib import Path

# Azure theme colors
COLORS = {
    'primary': (0, 120, 212),      # Azure blue
    'primary_dark': (16, 110, 190), # Darker azure
    'secondary': (108, 117, 125),   # Gray
    'success': (16, 124, 16),       # Green
    'white': (255, 255, 255),       # White
    'light_gray': (248, 249, 250),  # Light gray
    'dark_gray': (33, 37, 41),      # Dark gray
}

project_dir_path = Path(__file__).parent.parent

class LogoGenerator:
    """Generates logo and favicon for the Excel Column Mapper application"""
    
    def __init__(self, assets_dir = project_dir_path / "assets"):
        self.assets_dir = Path(assets_dir)
        self.assets_dir.mkdir(exist_ok=True)
        
    def create_logo(self, size=256):
        """Create the main application logo"""
        # Create transparent image
        img = Image.new('RGBA', (size, size), (255, 255, 255, 0))
        draw = ImageDraw.Draw(img)
        
        # Calculate proportions
        margin = size // 8
        grid_size = size - (2 * margin)
        cell_size = grid_size // 6
        
        # Draw background circle (optional)
        circle_radius = size // 2 - 10
        circle_center = size // 2
        # Uncomment for circular background:
        # draw.ellipse([10, 10, size-10, size-10], fill=COLORS['light_gray'] + (50,))
        
        # Draw spreadsheet grid
        grid_start_x = margin
        grid_start_y = margin + size // 6
        
        # Draw grid cells
        for row in range(4):
            for col in range(6):
                x1 = grid_start_x + col * cell_size
                y1 = grid_start_y + row * cell_size
                x2 = x1 + cell_size
                y2 = y1 + cell_size
                
                # Alternate cell colors for visual appeal
                if (row + col) % 2 == 0:
                    fill_color = COLORS['primary'] + (80,)
                else:
                    fill_color = COLORS['white'] + (120,)
                
                draw.rectangle([x1, y1, x2, y2], fill=fill_color, outline=COLORS['primary'], width=2)
        
        # Draw mapping arrows
        arrow_y = grid_start_y + (2 * cell_size) + cell_size // 2
        
        # Left arrow (source to mapping)
        arrow1_start = grid_start_x + (2 * cell_size)
        arrow1_end = arrow1_start + cell_size
        self.draw_arrow(draw, arrow1_start, arrow_y, arrow1_end, arrow_y, COLORS['success'], width=4)
        
        # Right arrow (mapping to destination)
        arrow2_start = grid_start_x + (4 * cell_size)
        arrow2_end = arrow2_start + cell_size
        self.draw_arrow(draw, arrow2_start, arrow_y, arrow2_end, arrow_y, COLORS['success'], width=4)
        
        # Draw title text
        try:
            # Try to use a nice font
            font_size = max(16, size // 16)
            font = ImageFont.truetype("arial.ttf", font_size)
        except (OSError, IOError):
            # Fallback to default font
            font = ImageFont.load_default()
        
        title_text = "EXCEL MAPPER"
        
        # Get text dimensions
        bbox = draw.textbbox((0, 0), title_text, font=font)
        text_width = bbox[2] - bbox[0]
        text_height = bbox[3] - bbox[1]
        
        # Center text at bottom
        text_x = (size - text_width) // 2
        text_y = size - margin - text_height
        
        # Draw text with shadow effect
        draw.text((text_x + 2, text_y + 2), title_text, fill=COLORS['dark_gray'] + (100,), font=font)
        draw.text((text_x, text_y), title_text, fill=COLORS['primary'], font=font)
        
        # Add version indicator (small dot)
        version_size = size // 32
        draw.ellipse([size - margin, margin, size - margin + version_size, margin + version_size], 
                    fill=COLORS['success'])
        
        return img
    
    def draw_arrow(self, draw, x1, y1, x2, y2, color, width=3):
        """Draw an arrow from (x1, y1) to (x2, y2)"""
        # Draw line
        draw.line([x1, y1, x2, y2], fill=color, width=width)
        
        # Draw arrowhead
        arrow_length = 8
        arrow_width = 6
        
        # Calculate arrowhead points
        if x2 > x1:  # Pointing right
            points = [
                (x2, y2),
                (x2 - arrow_length, y2 - arrow_width),
                (x2 - arrow_length, y2 + arrow_width)
            ]
        else:  # Pointing left
            points = [
                (x2, y2),
                (x2 + arrow_length, y2 - arrow_width),
                (x2 + arrow_length, y2 + arrow_width)
            ]
        
        draw.polygon(points, fill=color)
    
    def create_favicon(self, logo_img):
        """Create favicon from logo image"""
        # Create 32x32 favicon
        favicon_32 = logo_img.resize((32, 32), Image.Resampling.LANCZOS)
        
        # Create 16x16 favicon (simplified version)
        favicon_16 = Image.new('RGBA', (16, 16), (255, 255, 255, 0))
        draw = ImageDraw.Draw(favicon_16)
        
        # Simplified icon for 16x16 - just a grid with arrow
        # Draw mini grid
        for i in range(3):
            for j in range(3):
                x1, y1 = 2 + i * 4, 2 + j * 4
                x2, y2 = x1 + 3, y1 + 3
                
                if (i + j) % 2 == 0:
                    fill_color = COLORS['primary']
                else:
                    fill_color = COLORS['white']
                
                draw.rectangle([x1, y1, x2, y2], fill=fill_color, outline=COLORS['primary'])
        
        # Draw simple arrow
        draw.line([11, 8, 14, 8], fill=COLORS['success'], width=2)
        draw.polygon([(14, 8), (12, 6), (12, 10)], fill=COLORS['success'])
        
        return favicon_32, favicon_16
    
    def save_assets(self):
        """Generate and save all assets"""
        print("Generating Excel Column Mapper assets...")
        
        # Create main logo (256x256)
        print("Creating main logo (256x256)...")
        logo_256 = self.create_logo(256)
        logo_path = self.assets_dir / "logo.png"
        logo_256.save(logo_path)
        print(f"Logo saved: {logo_path}")
        
        # Create smaller logo versions
        print("Creating additional logo sizes...")
        sizes = [128, 64, 48]
        for size in sizes:
            logo_resized = logo_256.resize((size, size), Image.Resampling.LANCZOS)
            size_path = self.assets_dir / f"logo_{size}.png"
            logo_resized.save(size_path)
            print(f"Logo {size}x{size} saved: {size_path}")
        
        # Create favicons
        print("Creating favicons...")
        favicon_32, favicon_16 = self.create_favicon(logo_256)
        
        # Save individual favicon sizes
        favicon_32_path = self.assets_dir / "favicon_32.png"
        favicon_16_path = self.assets_dir / "favicon_16.png"
        favicon_32.save(favicon_32_path)
        favicon_16.save(favicon_16_path)
        
        # Create ICO file with multiple sizes
        ico_path = self.assets_dir / "favicon.ico"
        favicon_32.save(ico_path, format='ICO', sizes=[(32, 32), (16, 16)])
        print(f"Favicon ICO saved: {ico_path}")
        
        # Create app icon (for potential desktop integration)
        print("Creating app icon...")
        app_icon = logo_256.resize((128, 128), Image.Resampling.LANCZOS)
        app_icon_path = self.assets_dir / "app_icon.png"
        app_icon.save(app_icon_path)
        print(f"App icon saved: {app_icon_path}")
        
        print("\nAll assets generated successfully!")
        print("\nGenerated files:")
        print(f"{self.assets_dir}/")
        for file in sorted(self.assets_dir.glob("*")):
            if file.is_file():
                print(f"{file.name}")
    
    def create_sample_screenshots(self):
        """Create placeholder screenshots for documentation"""
        screenshots_dir = self.assets_dir / "screenshots"
        screenshots_dir.mkdir(exist_ok=True)
        
        # Create placeholder screenshots
        screenshot_size = (800, 600)
        screenshots = [
            ("main_interface.png", "Main Interface Screenshot"),
            ("mapping_process.png", "Column Mapping Process"),
            ("data_preview.png", "Data Preview")
        ]
        
        print("\nCreating placeholder screenshots...")
        
        for filename, title in screenshots:
            img = Image.new('RGB', screenshot_size, COLORS['light_gray'])
            draw = ImageDraw.Draw(img)
            
            # Draw border
            draw.rectangle([0, 0, screenshot_size[0]-1, screenshot_size[1]-1], 
                         outline=COLORS['primary'], width=3)
            
            # Add title
            try:
                font = ImageFont.truetype("arial.ttf", 24)
            except (OSError, IOError):
                font = ImageFont.load_default()
            
            bbox = draw.textbbox((0, 0), title, font=font)
            text_width = bbox[2] - bbox[0]
            text_x = (screenshot_size[0] - text_width) // 2
            text_y = screenshot_size[1] // 2
            
            draw.text((text_x, text_y), title, fill=COLORS['primary'], font=font)
            
            # Add placeholder note
            note = f"Replace with actual screenshot\n{filename}"
            try:
                small_font = ImageFont.truetype("arial.ttf", 14)
            except (OSError, IOError):
                small_font = ImageFont.load_default()
            
            note_bbox = draw.textbbox((0, 0), note, font=small_font)
            note_width = note_bbox[2] - note_bbox[0]
            note_x = (screenshot_size[0] - note_width) // 2
            note_y = text_y + 50
            
            draw.text((note_x, note_y), note, fill=COLORS['secondary'], font=small_font)
            
            # Save screenshot
            screenshot_path = screenshots_dir / filename
            img.save(screenshot_path)
            print(f"Screenshot placeholder saved: {screenshot_path}")


def main():
    """Main function to generate all assets and project files"""
    print("Excel Column Mapper - Asset Generator")
    print("=" * 50)
    
    # Generate visual assets
    generator = LogoGenerator()
    generator.save_assets()
    generator.create_sample_screenshots()
    
if __name__ == "__main__":
    main()
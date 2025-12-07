#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
创建 macOS Sequoia 风格的应用图标 - PPT Transfer
橙色到紫色渐变，文档提取主题
"""

from PIL import Image, ImageDraw
import os

def create_icon():
    # Icon size for macOS (1024x1024 for best quality)
    size = 1024

    # Create image with transparent background
    img = Image.new('RGBA', (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)

    # macOS Sequoia color palette - 橙色到紫色渐变
    accent_color_1 = (255, 149, 0)     # Apple Orange
    accent_color_2 = (175, 82, 222)    # Apple Purple

    # Draw rounded square background with gradient effect
    corner_radius = 180

    # Create gradient background
    for i in range(size):
        t = i / size
        r = int(accent_color_1[0] + (accent_color_2[0] - accent_color_1[0]) * t)
        g = int(accent_color_1[1] + (accent_color_2[1] - accent_color_1[1]) * t)
        b = int(accent_color_1[2] + (accent_color_2[2] - accent_color_1[2]) * t)
        draw.line([(0, i), (size, i)], fill=(r, g, b, 255))

    # Create a mask for rounded corners
    mask = Image.new('L', (size, size), 0)
    mask_draw = ImageDraw.Draw(mask)
    mask_draw.rounded_rectangle(
        [(0, 0), (size, size)],
        radius=corner_radius,
        fill=255
    )

    # Apply mask to create rounded corners
    output = Image.new('RGBA', (size, size), (0, 0, 0, 0))
    output.paste(img, (0, 0), mask)

    # Draw the document extraction icon
    draw = ImageDraw.Draw(output)

    # Center position
    center_x, center_y = size // 2, size // 2

    # Draw document icon (文档图标)
    doc_width = 280
    doc_height = 380
    doc_x = center_x - doc_width // 2
    doc_y = center_y - doc_height // 2 - 30

    # Draw document with shadow
    shadow_offset = 8
    draw.rounded_rectangle(
        [(doc_x + shadow_offset, doc_y + shadow_offset),
         (doc_x + doc_width + shadow_offset, doc_y + doc_height + shadow_offset)],
        radius=20,
        fill=(0, 0, 0, 60)
    )
    draw.rounded_rectangle(
        [(doc_x, doc_y), (doc_x + doc_width, doc_y + doc_height)],
        radius=20,
        fill=(255, 255, 255, 255),
        outline=(255, 255, 255, 100),
        width=3
    )

    # Draw lines inside document (文本线条)
    line_y_start = doc_y + 80
    line_spacing = 45
    for i in range(5):
        y = line_y_start + i * line_spacing
        line_width = doc_width - 60 if i % 2 == 0 else doc_width - 100
        draw.rounded_rectangle(
            [(doc_x + 30, y), (doc_x + line_width, y + 15)],
            radius=7,
            fill=(200, 200, 255, 180)
        )

    # Draw arrow pointing down and right (提取箭头)
    arrow_start_x = center_x + 80
    arrow_start_y = center_y + 50
    arrow_end_x = center_x + 180
    arrow_end_y = center_y + 150

    # Arrow shaft
    draw.line(
        [(arrow_start_x, arrow_start_y), (arrow_end_x, arrow_end_y)],
        fill=(255, 255, 255, 255),
        width=20
    )

    # Arrow head
    arrow_head_size = 35
    arrow_head = [
        (arrow_end_x, arrow_end_y),
        (arrow_end_x - arrow_head_size, arrow_end_y - arrow_head_size),
        (arrow_end_x - arrow_head_size - 10, arrow_end_y - arrow_head_size + 25)
    ]
    draw.polygon(arrow_head, fill=(255, 255, 255, 255))

    arrow_head2 = [
        (arrow_end_x, arrow_end_y),
        (arrow_end_x - arrow_head_size, arrow_end_y + arrow_head_size),
        (arrow_end_x - arrow_head_size + 25, arrow_end_y + arrow_head_size - 10)
    ]
    draw.polygon(arrow_head2, fill=(255, 255, 255, 255))

    # Draw extracted content icon (提取的内容)
    content_size = 100
    content_x = center_x + 170
    content_y = center_y + 160

    draw.rounded_rectangle(
        [(content_x, content_y), (content_x + content_size, content_y + content_size)],
        radius=15,
        fill=(255, 255, 255, 240),
        outline=(255, 255, 255, 150),
        width=2
    )

    # Draw "T" symbol inside (表示文本)
    draw.text(
        (content_x + content_size // 2, content_y + content_size // 2),
        "T",
        fill=(150, 100, 200, 255),
        font=None,
        anchor="mm"
    )

    # Save the icon in multiple sizes
    output.save('/Users/whitney/Downloads/PPT-Transfer/icon_1024.png', 'PNG')
    print("✓ Created icon_1024.png (1024x1024)")

    # Create standard macOS icon sizes
    sizes = [512, 256, 128, 64, 32, 16]
    for s in sizes:
        resized = output.resize((s, s), Image.Resampling.LANCZOS)
        resized.save(f'/Users/whitney/Downloads/PPT-Transfer/icon_{s}.png', 'PNG')
        print(f"✓ Created icon_{s}.png ({s}x{s})")

    print("\n✓ All icons created successfully!")
    print("Main icon: icon_1024.png")

if __name__ == '__main__':
    create_icon()

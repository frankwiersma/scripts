#!/usr/bin/env python3
from wand.image import Image
from wand.drawing import Drawing
from wand.color import Color
import os

# Timer settings
output_dir = "timer_gifs"
width, height = 1200, 800
font_size = 300

# Create output directory if it doesn't exist
os.makedirs(output_dir, exist_ok=True)

def nice_time(seconds):
    m, s = divmod(seconds, 60)
    return "{0:02d}:{1:02d}".format(m, s)

def create_timer_gif(start_minutes, start_seconds):
    filename = f"{output_dir}/timer_{start_minutes:02d}_{start_seconds:02d}.gif"
    total_seconds = start_minutes * 60 + start_seconds
    labels = [nice_time(s) for s in range(total_seconds, max(total_seconds - 60, -1), -1)]

    with Image() as gif:
        for label in labels:
            with Image(width=width, height=height, background=Color('transparent')) as frame:
                with Drawing() as draw:
                    draw.font = 'Arial'
                    draw.font_size = font_size
                    draw.text_alignment = 'center'
                    draw.text_antialias = True
                    draw.fill_color = Color("white")
                    
                    x = int(frame.width / 2)
                    y = int(frame.height / 2)
                    draw.text(x, y, label)
                    draw(frame)
                
                frame.delay = 100
                frame.dispose = 'background'
                gif.sequence.append(frame)
        
        gif.dispose = 'background'
        gif.delay = 100
        gif.loop = 0
        gif.type = 'optimize'
        gif.save(filename=filename)
    
    print(f"GIF saved as {filename}")

# Create 15 GIFs, from 15:00 to 01:00
for minutes in range(15, 0, -1):
    create_timer_gif(minutes, 0)

print("All timer GIFs have been created.")